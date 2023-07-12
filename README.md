Documentation for the Reconciliation Program


This program is designed to perform reconciliation tasks for financial transactions. 

It uses a graphical user interface (GUI) built with the Tkinter library in Python. The program can handle different types of reconciliation tasks, including EDC reconciliation, CP reconciliation, and CP transaction reconciliation.

Class: Reconciler
This class is responsible for creating the GUI and handling the reconciliation tasks.

Method: init
This method is the constructor for the Reconciler class. It initializes a new instance of the Reconciler class with a root Tkinter window. It also sets up the initial state of the GUI.

Method: choose_edc
This method is responsible for setting up the GUI for EDC reconciliation. It creates labels, buttons, and entry fields for the user to input the necessary information for EDC reconciliation. If the choose_edc_switch attribute is set to True, it sets up the GUI for EDC reconciliation. If it's set to False, it hides the GUI for EDC reconciliation.

Method: choose_cp_recon
This method is responsible for setting up the GUI for CP reconciliation. It creates labels, buttons, and entry fields for the user to input the necessary information for CP reconciliation. If the choose_cp_switch attribute is set to True, it sets up the GUI for CP reconciliation. If it's set to False, it hides the GUI for CP reconciliation.

Method: choose_cp_transaction
This method is responsible for setting up the GUI for CP transaction reconciliation. It creates labels, buttons, and entry fields for the user to input the necessary information for CP transaction reconciliation. If the choose_cp_trans attribute is set to True, it sets up the GUI for CP transaction reconciliation. If it's set to False, it hides the GUI for CP transaction reconciliation.

Method: run_reconciler_cp_trans
This method is responsible for performing the CP transaction reconciliation. It reads the CP transactions file, the settlement Excel file, and the authorization Excel file. It then performs a left join on the 'Auth Code' column to merge the data from the three files. It saves the results to an Excel file named "transaction_reconciliation_results_{merchant_name}.xlsx".

Method: run_reconciler_edc
This method is responsible for performing the EDC reconciliation. It reads the EDC PDF file, the settlement Excel file, and the authorization Excel file. It then performs a left join on the 'Auth Code' column to merge the data from the three files. It saves the results to an Excel file named "edc_reconciliation_results_{merchant_name}_{start_date}.xlsx".

Method: run_reconciler_cp
This method is responsible for performing the reconciliation process for batches and deposits. It reads the deposits Excel file, extracts deposit amounts and dates, and loads the instore batches PDF file. It then extracts batch amounts and dates from the instore batches PDF file. If an ecomm batches PDF file exists, it loads it and extracts batch amounts and dates from it. The method then combines the instore and ecomm batch amounts and dates into a single DataFrame. It groups by the same date if there's an ecomm batch and sums the amounts. If the is_capn attribute is set to True, it deducts AMEX amounts from batch amounts. The method then iterates over instore and ecomm dates to find matching date pairs. It creates a new column with the date shifted by one day, groups by consecutive pairs of days, and sums the batch amounts. It then merges the deposit DataFrame with the combined batch DataFrame on the Amount column. Finally, it saves the combined data to an Excel file named after the 'DBA Name' value.

Class: Tk
This is the main class for Tkinter. It is used to create a root window. This root window is used to manage all other widgets.

Method: init
This method is the constructor for the Tk class. It initializes a new instance of the Tk class.

Method: mainloop
This method is responsible for running the main event loop of the Tkinter application. It waits for user interactions and updates the GUI accordingly.

Function: main
This function creates an instance of the Tk class, an instance of the Reconciler class, and runs the main event loop of the Tkinter application.
