from tkinter import Tk, Button, Label
import pandas

def ChooseReport(on_report_selected):
    def report_selected(report_number):
        on_report_selected(report_number)

    # Create the main UI window
    root = Tk()
    root.title("Choose Report")

    # Add label
    label = Label(root, text="Select the report to export:")
    label.pack(pady=10)

    # Add buttons for each report
    button_report1 = Button(root, text="Growth Rate", command=lambda: report_selected(1))
    button_report1.pack(pady=5)

    button_report2 = Button(root, text="District Growth Per Year", command=lambda: report_selected(2))
    button_report2.pack(pady=5)

    button_report3 = Button(root, text="Volume Percentage", command=lambda: report_selected(3))
    button_report3.pack(pady=5)

    button_report4 = Button(root, text="E Class Market", command=lambda: report_selected(4))  
    button_report4.pack(pady=5)

    button_report5 = Button(root, text="D2 Class Market", command=lambda: report_selected(5))  
    button_report5.pack(pady=5)
    
    button_report5 = Button(root, text="D1 Class Market", command=lambda: report_selected(6))  
    button_report5.pack(pady=5)

    button_report5 = Button(root, text="B Class Market", command=lambda: report_selected(7))  
    button_report5.pack(pady=5)
    

    # Run the UI loop
    root.mainloop()
