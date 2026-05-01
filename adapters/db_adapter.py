import pyodbc
from datetime import datetime
from robot.api import logger
from robot.api.deco import keyword
from Resources.Config import Config

# ============================================================================
# DATABASE ADAPTER FOR RPA0032 - SIMPLIFIED & CLEAN
# ============================================================================

class db_adapter:
    """
    Simplified Database Adapter for RPA0032 Payment Processing
    - Handles SQL Server connections
    - Inserts payment records
    - Retrieves configuration values
    - Provides detailed logging of database operations
    """
    
    def __init__(self):
        """Initialize with connection string from Config.py in same directory"""
        self.connection_string = self._load_connection_string()
    
    def _load_connection_string(self):
        try:
            cfg = Config()
            CONNECTION_STRING = cfg.ConnectionString
            logger.console("✓ Connection string loaded from Config.py")
            logger.console(CONNECTION_STRING)
            return CONNECTION_STRING
        except ImportError as e:
            logger.warn(f"⚠ Config.py not found - using fallback connection string: {e}")
            fallback = (
                "Driver={SQL Server};Server=itlsqlotherscons;Database=UIPath_Param;Integrated Security=True"
            )
            logger.warn("⚠ Config.py not found - using fallback connection string")
            return fallback

    def _get_connection(self):
        """
        Establishes and returns database connection
        Returns None if connection fails
        """
        if not self.connection_string:
            logger.error("❌ No connection string available")
            return None
        
        try:
            conn = pyodbc.connect(self.connection_string)
            logger.info("✓ Database connection established")
            return conn
        except pyodbc.Error as e:
            logger.error(f"❌ Database connection failed: {e}")
            logger.error(f"Connection string: {self.connection_string}")
            return None

    # ========================================================================
    # MAIN OPERATIONS
    # ========================================================================

    @keyword("Insert And Update Payment Record")
    def insert_or_update_payment_record(self, data, vendor, CR_DR, status, barcodes):
        """
        Insert a new payment record into the tb_RPA0032_INDIA table or update an existing record 
        if the remitter_account and beneficiary_account already exist.

        Args:
            data (dict): Payment data with keys:
                remitter_name, remitter_address, remitter_account, 
                beneficiary_name, beneficiary_address, beneficiary_account, 
                value_date, company_code
            vendor (dict): Vendor data with keys:
                currency, amount, swift_name, swift_code, kz
            CR_DR (str): Transaction type (e.g., 'Outgoing')
            status (str): Payment status (e.g., 'Processed')
            barcodes (str): Barcode(s) for the transaction

        Returns:
            int: ID of the inserted or updated record, or None if insertion or update fails.
        """
        conn = self._get_connection()
        if not conn:
            return None
        try:
            cursor = conn.cursor()
            
            insert_query = """
                INSERT INTO [UIPath_Param].[dbo].[tb_RPA0032_INDIA] (
                    [status], [CR_DR], [currency], [amount], [remitter_name], 
                    [remitter_address], [remitter_account], [beneficiary_name], 
                    [beneficiary_address], [beneficiary_account], [swift_name], 
                    [swift_code], [value_date], [creation_date], [kz_number], 
                    [barcodes], [company_code] , [log_message]
                ) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            value_date=''
            raw_date = data.get('value_date', '')
            try:
                value_date = datetime.strptime(raw_date, "%d.%m.%Y").strftime("%Y-%m-%d")
            except ValueError:
                value_date = raw_date
            params = (
                "1",                                 # status (direct input)
                CR_DR,                                  # CR_DR (direct input)
                vendor.get('currency', ''),             # currency (from vendor)
                vendor.get('amount', 0),                # amount (from vendor)
                data.get('remitter_name', ''),          # remitter_name (from data)
                data.get('remitter_address', ''),       # remitter_address (from data)
                data.get('remitter_account', ''),       # remitter_account (from data)
                data.get('beneficiary_name', ''),       # beneficiary_name (from data)
                data.get('beneficiary_address', ''),    # beneficiary_address (from data)
                data.get('beneficiary_account', ''),    # beneficiary_account (from data)
                vendor.get('swift_name', ''),           # swift_name (from vendor)
                vendor.get('swift_code', ''),           # swift_code (from vendor)
                value_date,            # value_date (from data)
                datetime.now().strftime("%Y-%m-%d"),    # creation_date (current date)
                vendor.get('kz', ''),                   # kz_number (from vendor)
                barcodes,                               # barcodes (direct input)
                data.get('company_code', ''),          # company_code (from data)
                "Processed"                            # log_message (hardcoded)
            )
            
            cursor.execute(insert_query, params)
            conn.commit()
            
            # Get the ID of the inserted record
            cursor.execute("SELECT @@IDENTITY")
            inserted_id = cursor.fetchone()[0]
            
            logger.info(f"✓ Inserted record ID={inserted_id} for KZ={vendor.get('kz', 'N/A')}")
            # Return the ID of the inserted record

        except Exception as e:
            logger.error(f"Error inserting or updating payment record: {e}")
            return None
        finally:
            conn.close()

    @keyword("Get Latest Inserted Record")
    def get_latest_inserted_record(self):
        """
        Get the latest inserted record from the tb_RPA0032_INDIA table and print it.
        
        Args:
            connection_string (str): The connection string to connect to the SQL Server database.
        
        Returns:
            None: Prints the latest inserted record in the terminal.
        """
        try:
            # Connect to the database using the connection string
            conn = pyodbc.connect(self.connection_string )
            cursor = conn.cursor()

            # SQL query to fetch the most recent record based on Creation_Date
            query = """
                SELECT TOP 1 * 
                FROM [UIPath_Param].[dbo].[tb_RPA0032_INDIA] 
                ORDER BY [Creation_Date] DESC
            """
            
            cursor.execute(query)
            record = cursor.fetchone()
            logger.console(record)
            if record:
                # Get column headers (field names) for better alignment in the terminal
                column_names = [column[0] for column in cursor.description]
                
                # Print the headers
                print(" | ".join(column_names))
                print("-" * 100)
                
                # Print the values of the record
                print(" | ".join(str(value) if value is not None else "N/A" for value in record))
            else:
                print("No records found.")
            
            # Close the cursor and connection
            cursor.close()
            conn.close()
            
        except Exception as e:
            logger.error(f"Error fetching latest inserted record: {e}")
            print(f"Error: {e}")

    def _log_full_entry(self, cursor, record_id):
        """
        Fetch and log the complete inserted record
        
        Args:
            cursor: Database cursor
            record_id: ID of the record to fetch
        """
        try:
            query = """
                SELECT [ID_Line], [Status], [Unit], [CR_DR], [ValueDate], [Remitter], 
                       [RemitterAccNo], [Currency], [Amount], [Beneficiary], [BeneficiaryAccNo], 
                       [TransactionRefNo], [ContactName], [Tel], [Log_Message], [Creation_Date], 
                       [Barcodes], [Email], [Company], [CompanyCode], [KZ]
                FROM [UIPath_Param].[dbo].[tb_RPA0032_INDIA]
                WHERE [ID_Line] = ?
            """
            cursor.execute(query, (record_id,))
            row = cursor.fetchone()
            
            if row:
                logger.info("=" * 80)
                logger.info("FULL DATABASE ENTRY:")
                logger.info("=" * 80)
                logger.info(f"ID_Line:         {row[0]}")
                logger.info(f"Status:          {row[1]}")
                logger.info(f"Unit:            {row[2]}")
                logger.info(f"CR_DR:           {row[3]}")
                logger.info(f"ValueDate:       {row[4]}")
                logger.info(f"Remitter:        {row[5]}")
                logger.info(f"RemitterAccNo:   {row[6]}")
                logger.info(f"Currency:        {row[7]}")
                logger.info(f"Amount:          {row[8]}")
                logger.info(f"Beneficiary:     {row[9]}")
                logger.info(f"BeneficiaryAccNo:{row[10]}")
                logger.info(f"TransactionRefNo:{row[11]}")
                logger.info(f"ContactName:     {row[12]}")
                logger.info(f"Tel:             {row[13]}")
                logger.info(f"Log_Message:     {row[14]}")
                logger.info(f"Creation_Date:   {row[15]}")
                logger.info(f"Barcodes:        {row[16]}")
                logger.info(f"Email:           {row[17]}")
                logger.info(f"Company:         {row[18]}")
                logger.info(f"CompanyCode:     {row[19]}")
                logger.info(f"KZ:              {row[20]}")
                logger.info("=" * 80)
        except Exception as e:
            logger.error(f"❌ Could not fetch full entry: {e}")

    @keyword("Get Config Value")
    def get_config_value(self, config_name):
        """
        Get configuration value from database
        
        Args:
            config_name (str): Config name (e.g., 'SAP_User', 'PathTemp')
        
        Returns:
            str: Config value or None if not found
        
        Example:
            ${sap_user}= | Get Config Value | SAP_User |
        """
        conn = self._get_connection()
        if not conn:
            return None
        
        try:
            cursor = conn.cursor()
            query = """
                SELECT [Value] 
                FROM [UIPath_Param].[dbo].[RPA0032-APAC] 
                WHERE [Name] = ?
            """
            cursor.execute(query, (config_name,))
            result = cursor.fetchone()
            
            if result:
                logger.info(f"✓ Config '{config_name}' = '{result[0]}'")
                return result[0]
            else:
                logger.warn(f"⚠ Config '{config_name}' not found")
                return None
                
        except Exception as e:
            logger.error(f"❌ Failed to get config '{config_name}': {e}")
            return None
        finally:
            conn.close()

    @keyword("Load All Config Values")
    def load_all_config_values(self):
        """
        Load all configuration values as dictionary
        
        Returns:
            dict: All config name-value pairs
        
        Example:
            ${config}= | Load All Config Values |
        """
        conn = self._get_connection()
        if not conn:
            return {}
        
        try:
            cursor = conn.cursor()
            query = "SELECT [Name], [Value] FROM [UIPath_Param].[dbo].[RPA0032-APAC]"
            cursor.execute(query)
            
            config_dict = {row[0]: row[1] for row in cursor.fetchall()}
            logger.info(f"✓ Loaded {len(config_dict)} config values")
            return config_dict
            
        except Exception as e:
            logger.error(f"❌ Failed to load config: {e}")
            return {}
        finally:
            conn.close()

    @keyword("Check Payment Record Exists")
    def check_payment_record_exists(self, kz_number):
        """
        Check if payment record exists
        
        Args:
            kz_number (str): KZ transaction reference
        
        Returns:
            bool: True if exists, False otherwise
        
        Example:
            ${exists}= | Check Payment Record Exists | 12051464 |
        """
        conn = self._get_connection()
        if not conn:
            return False
        
        try:
            cursor = conn.cursor()
            query = "SELECT COUNT(*) FROM [UIPath_Param].[dbo].[tb_RPA0032_INDIA] WHERE [KZ] = ?"
            cursor.execute(query, (kz_number,))
            count = cursor.fetchone()[0]
            
            exists = count > 0
            logger.info(f"✓ Record exists for KZ {kz_number}: {exists}")
            return exists
            
        except Exception as e:
            logger.error(f"❌ Check failed for KZ {kz_number}: {e}")
            return False
        finally:
            conn.close()

    @keyword("Get Full Record By KZ")
    def get_full_record_by_kz(self, kz_number):
        """
        Retrieve complete record details by KZ number
        
        Args:
            kz_number (str): KZ transaction reference
        
        Returns:
            dict: Complete record data or None if not found
        
        Example:
            ${record}= | Get Full Record By KZ | 12051464 |
        """
        conn = self._get_connection()
        if not conn:
            return None
        
        try:
            cursor = conn.cursor()
            query = """
                SELECT [ID_Line], [Status], [Unit], [CR_DR], [ValueDate], [Remitter], 
                       [RemitterAccNo], [Currency], [Amount], [Beneficiary], [BeneficiaryAccNo], 
                       [TransactionRefNo], [Barcodes], [KZ], [Creation_Date], [Log_Message]
                FROM [UIPath_Param].[dbo].[tb_RPA0032_INDIA]
                WHERE [KZ] = ?
            """
            cursor.execute(query, (kz_number,))
            row = cursor.fetchone()
            
            if row:
                record = {
                    'id_line': row[0],
                    'status': row[1],
                    'unit': row[2],
                    'cr_dr': row[3],
                    'value_date': row[4],
                    'remitter': row[5],
                    'remitter_acc': row[6],
                    'currency': row[7],
                    'amount': row[8],
                    'beneficiary': row[9],
                    'iban': row[10],
                    'transaction_ref': row[11],
                    'barcodes': row[12],
                    'kz': row[13],
                    'creation_date': row[14],
                    'log_message': row[15]
                }
                logger.info(f"✓ Retrieved record ID_Line={record['id_line']} for KZ {kz_number}")
                return record
            else:
                logger.warn(f"⚠ No record found for KZ {kz_number}")
                return None
                
        except Exception as e:
            logger.error(f"❌ Failed to retrieve record for KZ {kz_number}: {e}")
            return None
        finally:
            conn.close()

    @keyword("Test Database Connection")
    def test_database_connection(self):
        """
        Test database connection
        
        Returns:
            bool: True if connected successfully
        
        Example:
            ${connected}= | Test Database Connection |
        """
        conn = self._get_connection()
        if not conn:
            return False
        
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            cursor.fetchone()
            logger.info("✓ Database connection test PASSED")
            return True
        except Exception as e:
            logger.error(f"❌ Connection test failed: {e}")
            return False
        finally:
            conn.close()
