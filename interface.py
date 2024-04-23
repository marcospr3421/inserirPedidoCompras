import tkinter as tk
from geraPedidoCompra import handle_soap_request


def call_handle_soap_request():
    handle_soap_request(handle_soap_request.url, handle_soap_request.headers, handle_soap_request.body)

# Create the GUI window
window = tk.Tk()

# Create a button to call the handle_soap_request function
button = tk.Button(window, text="Call handle_soap_request", command=call_handle_soap_request)
button.pack()

# Start the GUI event loop
window.mainloop()