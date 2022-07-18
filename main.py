import tkinter as tk
from tkinter import ttk
import pandas as pd
from datetime import date
import os
import re
import subprocess
import cx_Oracle
from sqlalchemy import create_engine

today = date.today()

class Layout(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry('495x330')
        self.resizable(width=False, height=False)
        self.configure(bg='#282a36', width=10, height=10)
        self.title("Query to Excel")
        self.create_frames()

    def create_frames(self):
        self.right_header = tk.Label(self, text="Select DB Environment",
                                     bg='#282a36', fg='#fff', font='Cascadia 10 bold')
        self.right_header.grid(row=0, column=0, pady=(10, 0), padx=10)

        self.left_header = tk.Label(self, text="Enter your query below",
                                    bg='#282a36', fg='#fff', font='Cascadia 10 bold')
        self.left_header.grid(row=0, column=1, pady=(10, 0), padx=10)

        self.user_query = tk.Text(self, height=10, width=30,
                                  fg='#ff5555', font='Cascadia 10 bold')
        self.user_query.grid(
            row=0, rowspan=7, column=1, pady=(30, 0), padx=30)

        self.save = tk.Label(self, text="Save file as:",
                             bg='#282a36', fg='#fff', font='Cascadia 10 bold')
        self.save.grid(row=7, column=1, pady=(0, 5))

        self.save_as = tk.Entry(self, width=30)
        self.save_as.grid(row=8, column=1, pady=(0, 10))

        self.selected_db = tk.IntVar()
        self.selected_db.set = '1'

        self.prd = tk.Radiobutton(self, text="PRD", variable=self.selected_db, value=1,
                                  bg='#282a36', fg='#fff', font='Cascadia 10 bold', selectcolor='#44475a')
        self.prd.grid(row=1, column=0, pady=(10, 3))

        self.stg = tk.Radiobutton(self, text="STG", variable=self.selected_db, value=2,  bg='#282a36', fg='#fff',
                                  font='Cascadia 10 bold', selectcolor='#44475a')
        self.stg.grid(row=2, column=0, pady=3)

        self.intdb = tk.Radiobutton(self, text="INT", variable=self.selected_db, value=3, bg='#282a36', fg='#fff',
                                    font='Cascadia 10 bold', selectcolor='#44475a')
        self.intdb.grid(row=3, column=0, pady=3)

        self.log_header = tk.Label(self, text="DB Login", bg='#282a36',
                                   fg='#fff', font='Cascadia 10 bold')
        self.log_header.grid(row=4, column=0, pady=(10, 2))

        self.pass_header = tk.Label(
            self, text="DB Password", bg='#282a36', fg='#fff', font='Cascadia 10 bold')
        self.pass_header.grid(
            row=6, column=0, pady=(10, 2))

        self.login = ttk.Entry(self)
        self.login.grid(row=5, column=0, padx=40)

        self.password = ttk.Entry(self, show='*')
        self.password.grid(row=7, column=0, padx=40)

        submit = tk.Button(self, text="Execute and Save", command=self.validate_vpn,
                           bg='#f55555', fg='#fff', font='Cascadia 10 bold')
        submit.grid(row=11, column=1, pady=(10, 0))

    def save_query(self, engine):
        query = self.user_query.get("1.0", "end")
        
        # Update path of where to save document
        writer = pd.ExcelWriter('path//queryresults_{today}.xlsx', datetime_format = 'YYYY-MMM-DD')
        df = pd.read_sql(query, con=engine)

        df.to_excel(
            f'writer', index=False) 


    def connection(self):
        pw = self.password.get()
        log = self.login.get()
        radio = self.selected_db.get()

        if radio == 1:
            radio = 'HOSTNAME//SERVICE NAME' # Update for access to Oracle DB
            host, service = radio.split('//')
        elif radio == 2:
            radio = 'HOSTNAME//SERVICE NAME'
            host, service = radio.split('//')
        elif radio == 3:
            radio = 'HOSTNAME//SERVICE NAME'
            host, service = radio.split('//')
        else:
            invalid_db = tk.Label(self, text="No DB Selected",
                                  bg='#282a36', fg='#fff', font='Cascadia 10 bold')
            invalid_db.grid(column=0, row=8, rowspan=9,
                            pady=(15, 5), padx=(5, 0))

        try:
            engine = (f'oracle+cx_oracle://{log}:{pw}@' +
                      cx_Oracle.makedsn(f'{host}', '1521',
                                        service_name=f'{service}')
                      )
            cx_Oracle.init_oracle_client(
                lib_dir=r"PATH\\instantclient_21_3") # You need to have Oracle Instant Client (https://www.oracle.com/database/technologies/instant-client.html)

            message = tk.Label(self, text="Querying DB",
                               bg='#282a36', fg='#fff', font='Cascadia 10 bold')
            message.grid(column=0, row=8, rowspan=9, pady=(15, 5), padx=(5, 0))

            self.save_query(engine)
        except:
            error = tk.Label(self, text="You entered an invalid query.",
                             bg='#282a36', fg='#fff', font='Cascadia 10 bold')
            error.grid(column=0, row=8, rowspan=9, pady=(15, 5), padx=(5, 0))

    def validate_vpn(self):
        '''Validate whether or not the user is connect to VPN. If not, display message and exit program.'''

        addresses = os.popen(
            'IPCONFIG | FINDSTR /R "Ethernet adapter Local Area Connection .* Address.*[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*"')

        first_eth_address = re.search(
            r'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', addresses.read()).group()

        host = first_eth_address
        ping = subprocess.Popen(["ping.exe", "-n", "1", "-w",
                                 "1", host], stdout=subprocess.PIPE).communicate()[0]

        if ('unreachable' in str(ping)) or ('timed' in str(ping)) or ('failure' in str(ping)):
            ping_chk = 0
        else:
            ping_chk = 1

        if ping_chk == 1 and host.startswith('IP ADDDRESS'): # update the IP address here
            self.connection()
        else:
            message = tk.Label(self, text="VPN Not Connected",
                               bg='#282a36', fg='#fff', font='Cascadia 10 bold')
            message.grid(column=0, row=8, rowspan=9, pady=(15, 5), padx=(5, 0))


def main():
    win = Layout()
    win.mainloop()


if __name__ == '__main__':
    main()
