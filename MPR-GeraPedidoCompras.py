import logging
from turtle import color
import requests
from io import StringIO
import pandas as pd
from typing import List, Tuple
from collections import defaultdict
from tkinter import Tk, filedialog, font
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox
import xml.sax.saxutils as saxutils
import os
from datetime import datetime

def gera_pedido_compra(pedidos: List[dict]) -> str:
    """
    Generates an XML payload containing the purchase order data.

    Args:
        pedidos: A list of dictionaries, each representing a purchase order.

    Returns:
        The XML payload as a string.
    """

    pedidos_compra = [
        f"""
        <PedidoCompra>
            <NumeroPedido>{saxutils.escape(pedido['numeroPedido'])}</NumeroPedido>
            <CnpjFornecedor>{saxutils.escape(pedido['cnpjFornecedor'])}</CnpjFornecedor>
            <Vendedor>{saxutils.escape(pedido['vendedor'])}</Vendedor>
            <ValorFrete>{saxutils.escape(pedido['valorFrete'])}</ValorFrete>
            <ValorImpostos>{saxutils.escape(pedido['valorImpostos'])}</ValorImpostos>
            <TipoPedido>{saxutils.escape(pedido['tipoPedido'])}</TipoPedido>
            <TipoPrazoPagamento>{saxutils.escape(pedido['tipoPrazoPagamento'])}</TipoPrazoPagamento>
            <PrazosPagamento>{saxutils.escape(pedido['prazosPagamento'])}</PrazosPagamento>
            <CodigoProduto>{saxutils.escape(pedido['codigoProduto'])}</CodigoProduto>
            <Quantidade>{saxutils.escape(pedido['quantidade'])}</Quantidade>
            <PrecoBruto>{saxutils.escape(pedido['precoBruto'])}</PrecoBruto>
            <DescontoTotal>{saxutils.escape(pedido['descontoTotal'])}</DescontoTotal>
            <AliquotaICMS>{saxutils.escape(pedido['aliquotaICMS'])}</AliquotaICMS>
            <AliquotaIPI>{saxutils.escape(pedido['aliquotaIPI'])}</AliquotaIPI>
            <PrevisaoEntrega>{saxutils.escape(pedido['previsaoEntrega'])}</PrevisaoEntrega>
            <ConcluirPedido>{saxutils.escape(pedido['concluirPedido'])}</ConcluirPedido>
            <UnidadeNegocio>{saxutils.escape(pedido['unidadeNegocio'])}</UnidadeNegocio>
            <NumeroPedidoVenda>{saxutils.escape(pedido['numeroPedidoVenda'])}</NumeroPedidoVenda>
        </PedidoCompra>"""
        for pedido in pedidos
    ]
    return "\n".join(pedidos_compra)

def select_file() -> str:
    """
    Prompts the user to select an Excel file and returns the file path.

    Returns:
        The file path of the selected Excel file.
    """

    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filename:
        print(f"File selected: {filename}")
    else:
        print("No file selected")
    return filename

def set_order(filename: str) -> List[Tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]]:
    """
    Reads the Excel file and returns a list of purchase order data.

    Args:
        filename: The path to the Excel file.

    Returns:
        A list of tuples, each representing a purchase order.
    """

    try:
        workbook = pd.read_excel(filename)
        rows = []
        for index, row in workbook.iterrows():
            rows.append(tuple(str(row.iloc[i]) for i in range(18)))
        return rows
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        return []


def handle_soap_request(url: str, headers: dict, body: str) -> List[Tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]]:
    """
    Sends the SOAP request and returns the parsed response.

    Args:
        url: The URL of the SOAP endpoint.
        headers: The HTTP headers for the request.
        body: The XML payload for the request.

    Returns:
        A list of tuples, each representing a purchase order response.
    """

    try:
        response = requests.post(url, data=body, headers=headers)
        response_content = response.content.decode('utf-8')
        response_content = ''.join(char for char in response_content if ord(char) < 255)

        xml_io = StringIO(response_content)
        result = []
        values = []

        for event, element in ET.iterparse(xml_io, events=("start", "end")):
            if event == "start" and element.tag == '{http://www.kplsolucoes.com.br/ABACOSWebService}InserirPedidoCompra':
                values = tuple(
                    element.findtext(f".//{{http://www.kplsolucoes.com.br/ABACOSWebService}}{tag}")
                    for tag in [
                        "NumeroPedido",
                        "CnpjFornecedor",
                        "Vendedor",
                        "ValorFrete",
                        "ValorImpostos",
                        "TipoPedido",
                        "TipoPrazoPagamento",
                        "PrazosPagamento",
                        "CodigoProduto",
                        "Quantidade",
                        "PrecoBruto",
                        "DescontoTotal",
                        "AliquotaICMS",
                        "AliquotaIPI",
                        "PrevisaoEntrega",
                        "ConcluirPedido",
                        "UnidadeNegocio",
                        "NumeroPedidoVenda"
                    ]
                )

                if None not in values:
                    result.append(values)

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")

    return result

def save_request_to_file(numero_pedido: str, request_body: str, response_content: str) -> str:
    """
    Saves the request and response details to a txt file.

    Args:
        numero_pedido: The purchase order number
        request_body: The XML request body
        response_content: The response content from the server

    Returns:
        The path to the generated file
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"pedido_{numero_pedido}_{timestamp}.txt"
    
    # Create logs directory if it doesn't exist
    os.makedirs("logs", exist_ok=True)
    
    filepath = os.path.join("logs", filename)
    
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"=== Pedido de Compra {numero_pedido} ===\n")
        f.write(f"Data/Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("=== REQUEST ===\n")
        f.write(request_body)
        f.write("\n\n=== RESPONSE ===\n")
        f.write(response_content)
    
    return filepath

def main():
    """
    Main function of the script.
    """

    # Create the main window
    root = tk.Tk()
    root.title("MPRLabs - GeraPedidosCompras - v1.0.10.25")
    root.geometry("700x500")
    root.resizable(False, False)
    root.configure(background="#ffffff")
    
    try:
        image_path = "mprLabs4sml.png"  # Replace with the actual path to your image file
        image = tk.PhotoImage(file=image_path)
        label = tk.Label(root, image=image)
        label.place(x=10, y=10, relwidth=1, relheight=1)
    except Exception as e:
        logging.warning("Failed to load background image: %s", e)
        
    
    
    try:
        icon_path = "mprIco.ico"  # Replace with the actual path to your icon file
        root.iconbitmap(icon_path)
    except Exception as e:
        logging.warning("Failed to set window icon: %s", e)


 
    # Create a label for the status
    status_label = tk.Label(root, text="")
    status_label.pack()

    # Create a progress bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
    progress_bar.pack()

    # Create a button to start the process
    start_button = tk.Button(root, text="Iniciar processamento", command=lambda: start_process(status_label, progress_bar, root), bg="green", fg="white")
    start_button.pack(side="bottom", padx=10, pady=5)

    
    
    # Create a button to exit the program
    exit_button = tk.Button(root, text="Sair", command=root.destroy, bg="red", fg="white")
    exit_button.pack(side="bottom")



    # Run the main loop
    root.mainloop()

def start_process(status_label, progress_bar, root):
    """
    Starts the process of generating and sending purchase orders.

    Args:
        status_label: The label to display status messages.
        progress_bar: The progress bar to update.
    """

    # Select the Excel file
    filename = select_file()
    status_label.config(text="Arquivo: " + filename)
    status_label.config(font=("Arial", 9, "bold"))
    

    if not filename:
        status_label.config(text="Nenhum arquivo selecionado.")
        return

    # Display a confirmation message
    result = messagebox.askquestion("Confirmação", "Deseja confirmar o processamento dos pedidos?")

    if result == "yes":
        status_label.config(text="Processamento confirmado.")

        # Set the purchase order data
        rows = set_order(filename)
        if not rows:
            status_label.config(text="Erro ao ler o arquivo Excel.")
            return

        # Group rows by purchase order number
        grouped_rows = defaultdict(list)
        for row in rows:
            grouped_rows[row[0]].append(row)

        # Process each purchase order
        total_orders = len(grouped_rows)
        current_order = 1
        
        # Create a list to store all generated file paths
        generated_files = []
        
        for numero_pedido, grouped_row in grouped_rows.items():
            pedidos = []
            
            for row in grouped_row:
                numero_pedido, cnpj_fornecedor, vendedor, valor_frete, valor_impostos, tipo_pedido, tipo_prazo_pagamento, prazos_pagamento, codigo_produto, quantidade, preco_bruto, desconto_total, aliquota_icms, aliquota_ipi, previsao_entrega, concluir_pedido, unidade_negocio, numero_pedido_venda = row
                pedidos.append({
                    'numeroPedido': numero_pedido,
                    'cnpjFornecedor': cnpj_fornecedor,
                    'vendedor': vendedor,
                    'valorFrete': valor_frete,
                    'valorImpostos': valor_impostos,
                    'tipoPedido': tipo_pedido,
                    'tipoPrazoPagamento': tipo_prazo_pagamento,
                    'prazosPagamento': prazos_pagamento,
                    'codigoProduto': codigo_produto,
                    'quantidade': quantidade,
                    'precoBruto': preco_bruto,
                    'descontoTotal': desconto_total,
                    'aliquotaICMS': aliquota_icms,
                    'aliquotaIPI': aliquota_ipi,
                    'previsaoEntrega': previsao_entrega,
                    'concluirPedido': concluir_pedido,
                    'unidadeNegocio': unidade_negocio,
                    'numeroPedidoVenda': numero_pedido_venda
                })

            # Generate the XML payload
            body = f"""<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <InserirPedidoCompra xmlns="http://www.kplsolucoes.com.br/ABACOSWebService">
                        <ChaveIdentificacao>E19C1DD8-345F-4F1B-A53F-CA4312CAF457</ChaveIdentificacao>
                        <ListaDePedidosCompra>
                            {gera_pedido_compra(pedidos)}
                        </ListaDePedidosCompra>
                    </InserirPedidoCompra>
                </soap:Body>
            </soap:Envelope>"""

            # Send the SOAP request and handle the response
            url = "http://ws.ocikpl.onclick.com.br/AbacosWSerp.asmx"
            headers = {'content-type': 'text/xml'}
            response = requests.post(url=url, headers=headers, data=body)
            response_content = response.content.decode('utf-8')
            
            # Save request and response to file
            filepath = save_request_to_file(numero_pedido, body, response_content)
            generated_files.append(filepath)
            
            xml_io = StringIO(response_content)
            for event, element in ET.iterparse(xml_io, events=("start", "end")):
                if event == "start" and element.tag == '{http://www.kplsolucoes.com.br/ABACOSWebService}InserirPedidoCompraResult':
                    codigo = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Codigo")
                    descricao = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Descricao")
                    tipo = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Tipo")
                    exceptionMessage = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}ExceptionMessage")
                    status_label.config(text=f"Código: {codigo}\nDescrição: {descricao}\nTipo: {tipo}\nMensagem de exceção: {exceptionMessage}")

            # Update the progress bar
            progress_bar["value"] = (current_order / total_orders) * 100
            root.update_idletasks()
            current_order += 1

        # Show summary of generated files
        files_message = f"Arquivos gerados na pasta 'logs':\n" + "\n".join(generated_files)
        messagebox.showinfo("Arquivos Gerados", files_message)

    else:
        status_label.config(text="Processamento cancelado.")
        
    # Show a message when finished
    messagebox.showinfo("Concluído", "O processamento foi concluído.")

if __name__ == "__main__":
    main()
