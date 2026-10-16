#!/usr/bin/env python3
import argparse
import json
import os
import re

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except ImportError:
    try:
        import psycopg2_binary as psycopg2
        from psycopg2.extras import RealDictCursor
    except ImportError:
        print("Please install psycopg2 or psycopg2-binary to run this script.")
        exit(1)


def load_env_file(env_file_path: str = ".env"):
    if not os.path.exists(env_file_path):
        env_file_path = os.path.join(os.path.dirname(__file__), env_file_path)
    if os.path.exists(env_file_path):
        with open(env_file_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, value = line.split("=", 1)
                    os.environ[key.strip()] = value.strip().strip("'").strip('"')


def main():
    parser = argparse.ArgumentParser(description="Generate SQL update script from markdown validation summary")
    parser.add_argument("--file", required=True, help="Path to validation_summary_*.md")
    parser.add_argument("--env-file", default=".env", help="Path to .env file")
    parser.add_argument("--out-dir", default="", help="Output directory for the generated SQL scripts (defaults to same folder as markdown file)")

    args = parser.parse_args()

    load_env_file(args.env_file)

    agency_id = int(os.environ.get("AGENCY_ID", "5"))

    try:
        db_url = os.environ.get("DATABASE_URL")
        db_user = os.environ.get("DB_USER")
        db_password = os.environ.get("DB_PASSWORD")

        if db_url:
            clean_url = db_url.strip()
            if clean_url.lower().startswith("jdbc:"):
                clean_url = clean_url[5:]
            conn_kwargs = {"dsn": clean_url, "connect_timeout": 5}
            if db_user:
                conn_kwargs["user"] = db_user
            if db_password:
                conn_kwargs["password"] = db_password
            print("🔌 Connecting to PostgreSQL database...", flush=True)
            db_conn = psycopg2.connect(**conn_kwargs)
        else:
            print("🔌 Connecting to PostgreSQL database...", flush=True)
            db_conn = psycopg2.connect(
                host=os.environ.get("DB_HOST", "localhost"),
                port=os.environ.get("DB_PORT", "5432"),
                dbname=os.environ.get("DB_NAME", "postgres"),
                user=db_user or "postgres",
                password=db_password or "",
                connect_timeout=5
            )
        db_conn.autocommit = True
        cursor = db_conn.cursor(cursor_factory=RealDictCursor)
        print("🟢 Connected to PostgreSQL database successfully.", flush=True)
    except Exception as e:
        print(f"❌ Failed to connect to database: {e}", flush=True)
        return

    with open(args.file, "r") as f:
        lines = f.readlines()

    mismatches = []
    current_policy = None
    in_table = False

    for line in lines:
        line = line.strip()
        if line.startswith("**Policy No:**"):
            # Extract Policy No
            m = re.search(r"\*\*Policy No:\*\*\s*(.*?)\s*\|", line)
            if m:
                current_policy = m.group(1).strip()
            in_table = False
        elif line.startswith("| Order |"):
            in_table = True
        elif in_table and line.startswith("|---"):
            continue
        elif in_table and line.startswith("|"):
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 7:
                # Strip markdown bolding and emojis added for highlighting
                order_str = parts[1].replace("*", "").replace("❌", "").replace("⚠️", "").strip()
                file_agent_str = parts[3].replace("*", "").replace("❌", "").replace("⚠️", "").strip()
                db_agent_str = parts[4].replace("*", "").replace("❌", "").replace("⚠️", "").strip()
                file_value_str = parts[5].replace("*", "").replace("❌", "").replace("⚠️", "").replace("%", "").strip()

                try:
                    order = int(order_str)
                except ValueError:
                    continue

                file_val = None
                try:
                    if file_value_str and file_value_str != "_N/A_":
                        file_val = float(file_value_str)
                except ValueError:
                    pass

                if file_agent_str and file_agent_str != "_N/A_":
                    # Parse File Agent Name and Code. Format: NAME (CODE)
                    m_agent = re.match(r"^(.*?)\s*\(([^)]+)\)$", file_agent_str)
                    if m_agent:
                        agent_name = m_agent.group(1).strip()
                        agent_code = m_agent.group(2).strip()
                        mismatches.append({
                            "policy_number": current_policy,
                            "order": order,
                            "agent_name": agent_name,
                            "agent_code": agent_code,
                            "raw_db_agent": db_agent_str,
                            "file_value": file_val
                        })
        elif in_table and not line:
            in_table = False

    # Group mismatches by policy
    policy_mismatches = {}
    for m in mismatches:
        policy_mismatches.setdefault(m["policy_number"], []).append(m)

    print(f"Found {len(mismatches)} mismatches across {len(policy_mismatches)} unique policies in the markdown file.")
    print("Connecting to database and generating updates...")

    import copy

    out_lines_uuid = []
    out_lines_value = []
    out_lines_both = []
    select_uuids = set()
    processed_count = 0
    
    # Caches for DB lookups
    user_cache = {}
    agency_cache = {}

    for pol_num, orders_to_fix in policy_mismatches.items():
        processed_count += 1
        if processed_count % 50 == 0:
            print(f"Processed {processed_count} / {len(policy_mismatches)} policies...")

        # Fetch policy_commission
        cursor.execute("""
            SELECT pc.uuid, pc.commission_split 
            FROM policies p
            JOIN policy_commission pc ON pc.policy_uuid = p.uuid
            WHERE p.number = %s AND p.agency_id = %s
        """, (pol_num, agency_id))

        rows = cursor.fetchall()
        if not rows:
            print(f"Warning: Policy {pol_num} not found in DB.")
            continue

        for row in rows:
            pc_uuid = row["uuid"]
            commission_split = row["commission_split"]

            if not isinstance(commission_split, list):
                continue

            split_uuid = copy.deepcopy(commission_split)
            split_value = copy.deepcopy(commission_split)
            split_both = copy.deepcopy(commission_split)

            updated_uuid = False
            updated_value = False
            updated_both = False

            for fix in orders_to_fix:
                order = fix["order"]
                agent_name = fix["agent_name"]
                agent_code = fix["agent_code"]
                file_val = fix["file_value"]
                new_uuid = None

                # Rule evaluations
                if order == 1:
                    normalized_name = re.sub(r'\s+', ' ', agent_name).strip().upper()
                    if normalized_name in ["MAIN LINE BENEFITS", "MLB", "MLB PARALLEL RUN"]:
                        pass
                    else:
                        if agency_id in agency_cache:
                            new_uuid = agency_cache[agency_id]
                        else:
                            cursor.execute("SELECT uuid FROM agencies WHERE id = %s", (agency_id,))
                            arow = cursor.fetchone()
                            if arow:
                                new_uuid = arow["uuid"]
                                agency_cache[agency_id] = new_uuid
                            else:
                                print(f"Warning: Agency ID {agency_id} not found.")
                else:
                    if agent_code in user_cache:
                        new_uuid = user_cache[agent_code]
                    else:
                        cursor.execute("SELECT uuid FROM users WHERE code = %s AND tenant_id = %s LIMIT 1", (agent_code, agency_id))
                        urow = cursor.fetchone()
                        if urow:
                            new_uuid = urow["uuid"]
                            user_cache[agent_code] = new_uuid
                        else:
                            print(f"Warning: User with code {agent_code} not found for policy {pol_num}.")

                if not new_uuid and file_val is None:
                    continue

                # Update the JSON arrays
                for item in split_uuid:
                    if item.get("order") == order:
                        if new_uuid and (item.get("user_id") != str(new_uuid) or item.get("user_code") != agent_code):
                            item["user_id"] = str(new_uuid)
                            item["user_name"] = agent_name
                            item["user_code"] = agent_code
                            updated_uuid = True
                        break

                for item in split_value:
                    if item.get("order") == order:
                        if file_val is not None and item.get("commission_value") != file_val:
                            item["commission_value"] = file_val
                            updated_value = True
                        break

                for item in split_both:
                    if item.get("order") == order:
                        if new_uuid and (item.get("user_id") != str(new_uuid) or item.get("user_code") != agent_code):
                            item["user_id"] = str(new_uuid)
                            item["user_name"] = agent_name
                            item["user_code"] = agent_code
                            updated_both = True
                        if file_val is not None and item.get("commission_value") != file_val:
                            item["commission_value"] = file_val
                            updated_both = True
                        break

            if updated_uuid or updated_value or updated_both:
                select_uuids.add(pc_uuid)

            if updated_uuid:
                out_lines_uuid.append(f"-- Fixing UUIDs for Policy {pol_num} (Commission Rule UUID: {pc_uuid})")
                out_lines_uuid.append(f"UPDATE policy_commission SET commission_split = '{json.dumps(split_uuid).replace(chr(39), chr(39)+chr(39))}'::jsonb WHERE uuid = '{pc_uuid}';\n")
            if updated_value:
                out_lines_value.append(f"-- Fixing Values for Policy {pol_num} (Commission Rule UUID: {pc_uuid})")
                out_lines_value.append(f"UPDATE policy_commission SET commission_split = '{json.dumps(split_value).replace(chr(39), chr(39)+chr(39))}'::jsonb WHERE uuid = '{pc_uuid}';\n")
            if updated_both:
                out_lines_both.append(f"-- Fixing Both for Policy {pol_num} (Commission Rule UUID: {pc_uuid})")
                out_lines_both.append(f"UPDATE policy_commission SET commission_split = '{json.dumps(split_both).replace(chr(39), chr(39)+chr(39))}'::jsonb WHERE uuid = '{pc_uuid}';\n")


    out_dir = args.out_dir if args.out_dir else (os.path.dirname(args.file) if os.path.dirname(args.file) else ".")
    os.makedirs(out_dir, exist_ok=True)
    
    path_uuid = os.path.join(out_dir, "update_uuid.sql")
    with open(path_uuid, "w") as f:
        f.write("\n".join(out_lines_uuid))

    path_value = os.path.join(out_dir, "update_value.sql")
    with open(path_value, "w") as f:
        f.write("\n".join(out_lines_value))

    path_both = os.path.join(out_dir, "update_both.sql")
    with open(path_both, "w") as f:
        f.write("\n".join(out_lines_both))

    out_lines_select = []
    if select_uuids:
        out_lines_select.append("SELECT * FROM policy_commission WHERE uuid IN (")
        uuids_list = sorted(list(select_uuids))
        for i, uid in enumerate(uuids_list):
            comma = "," if i < len(uuids_list) - 1 else ""
            out_lines_select.append(f"    '{uid}'{comma}")
        out_lines_select.append(");\n")

    path_select = os.path.join(out_dir, "select_all.sql")
    with open(path_select, "w") as f:
        f.write("\n".join(out_lines_select))

    print(f"Successfully generated {len(out_lines_select)//2} SELECT statements in {path_select}")
    print(f"Successfully generated {len(out_lines_uuid)//2} UUID update statements in {path_uuid}")
    print(f"Successfully generated {len(out_lines_value)//2} Value update statements in {path_value}")
    print(f"Successfully generated {len(out_lines_both)//2} Combined update statements in {path_both}")


if __name__ == "__main__":
    main()
