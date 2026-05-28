def update_prescriptions_table():
    """Update prescriptions table to allow NULL patient_id"""
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Check if prescriptions table exists
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='prescriptions'")
        if cur.fetchone():
            # Get current table info
            cur.execute("PRAGMA table_info(prescriptions)")
            columns = [col[1] for col in cur.fetchall()]
            
            # Check if patient_id column is set to NOT NULL
            cur.execute("PRAGMA table_info(prescriptions)")
            for col in cur.fetchall():
                if col[1] == 'patient_id' and col[3] == 1:  # notnull = 1
                    # Need to recreate table without NOT NULL constraint
                    print("Updating prescriptions table to allow NULL patient_id...")
                    
                    # Create temporary table
                    cur.execute("""
                        CREATE TABLE prescriptions_temp (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            prescription_no VARCHAR(50) UNIQUE NOT NULL,
                            patient_id INTEGER,
                            patient_name VARCHAR(200) NOT NULL,
                            doctor_id INTEGER NOT NULL,
                            doctor_name VARCHAR(100) NOT NULL,
                            drug_id INTEGER,
                            drug_name VARCHAR(100) NOT NULL,
                            strength VARCHAR(50),
                            dosage VARCHAR(100) NOT NULL,
                            frequency VARCHAR(100) NOT NULL,
                            duration VARCHAR(50) NOT NULL,
                            quantity INTEGER NOT NULL,
                            instructions TEXT,
                            refills INTEGER DEFAULT 0,
                            is_controlled BOOLEAN DEFAULT 0,
                            status VARCHAR(20) DEFAULT 'Active',
                            prescribed_date DATE NOT NULL,
                            prescribed_time TIME NOT NULL,
                            expires_date DATE,
                            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                        )
                    """)
                    
                    # Copy data
                    cur.execute("""
                        INSERT INTO prescriptions_temp SELECT * FROM prescriptions
                    """)
                    
                    # Drop old table
                    cur.execute("DROP TABLE prescriptions")
                    
                    # Rename temp table
                    cur.execute("ALTER TABLE prescriptions_temp RENAME TO prescriptions")
                    
                    conn.commit()
                    print("Prescriptions table updated successfully!")
                    break
            
    except Exception as e:
        app.logger.error(f"Error updating prescriptions table: {e}")
    finally:
        cur.close()
        conn.close()