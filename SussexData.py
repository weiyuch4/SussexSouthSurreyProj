import datetime
import pyodbc
import os


class SussexData:

    def __init__(self):
        os.chdir("..")
        parent_dir_path = os.path.abspath(os.curdir)
        self.conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            rf'DBQ={parent_dir_path}\South Surrey.mdb;'
        )
        self.cur = self.conn.cursor()
        os.chdir(f"{parent_dir_path}\RENEWAL")

    def generate_mail_list(self):
        try:
            start_date_input_str = input("\nEnter Start Date (DD/MM/YYYY): ")
            start_date = datetime.datetime.strptime(start_date_input_str, '%d/%m/%Y')

            end_date_input_str = input("\nEnter End Date (DD/MM/YYYY): ")
            end_date = datetime.datetime.strptime(end_date_input_str, '%d/%m/%Y')

            if start_date <= end_date:
                query = (
                    f"""select c.lastname, c.firstname, c.address, c.city, c.postalcode, 
                    pd.expiry, pd.year, pd.made, pd.model, pn.plate, pn.rodl, c.notes 
                    from (clients c 
                    inner join [policy details] pd on c.policyno = pd.policyno) 
                    inner join [policy no] pn on pd.policyno = pn.policyno 
                    where (pd.expiry between #{start_date}# and #{end_date}#) and c.notification = 'Mail' 
                    and pn.plate is not null;"""
                )
                self.cur.execute(query)
                mail = self.cur.fetchall()
                print("\nCount: " + str(len(mail)))
                return mail
            else:
                raise Exception("Start date must be earlier than end date")

        except Exception as err:
            print(f"Unexpected {err=}, {type(err)=}")

