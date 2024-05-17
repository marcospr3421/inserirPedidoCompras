
import logging
import requests
from io import StringIO
import pandas as pd
from typing import List, Tuple
from collections import defaultdict
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
import xml.sax.saxutils as saxutils

def gera_pedido_compra(pedidos: List[dict]) -> str:
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
        </PedidoCompra>"""
        for pedido in pedidos
    ]
    return "\n".join(pedidos_compra)
    


    

def read_excel_file(filename):
    """
    Reads an Excel file and returns the data as a list of tuples.

    Args:
        filename (str): The path to the Excel file.

    Returns:
        list: The data from the Excel file as a list of tuples.
    """
    try:
        workbook = pd.read_excel(filename)
        workbook = workbook.astype(str)
        rows = []
        for index, row in workbook.iterrows():
            rows.append(tuple(row))
        return rows
    except FileNotFoundError as e:
        logging.error("File not found: %s", e)
        raise e
    except IsADirectoryError as e:
        logging.error("Is a directory: %s", e)
        raise e
    except PermissionError as e:
        logging.error("Permission denied: %s", e)
        raise e
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        raise e
def select_file():
    global filename
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filename:
        print(f"File selected: {filename}")
    else:
        print("No file selected")

window = tk.Tk()
window.title("File Selector")

filename = ''  # Initialize filename as an empty string
button1 = tk.Button(window, text="Browse", command=select_file)
button1.pack(side=tk.LEFT)

button2 = tk.Button(window, text="Import", command=lambda: read_excel_file(filename) if filename else print("No file selected"))
button2.pack(side=tk.LEFT)

window.mainloop()


def setOrder() -> List[Tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]]:

    try:
        


        workbook = pd.read_excel(filename)
        rows = []
        for index, row in workbook.iterrows():
            rows.append(tuple(str(row.iloc[i]) for i in range(16)))
        return rows
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)

        
        return []
rows = setOrder()
# Group rows by numeroPedido
grouped_rows = defaultdict(list)
for row in rows:
    grouped_rows[row[0]].append(row)
    
results = []







def handle_soap_request(url: str, headers: dict, body: str) -> List[Tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]]:
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
                        "ConcluirPedido"
                        
                    ]
                )

                        
        if None not in values:
            return result
        result.append(values)
        
                
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        
    
    return []  # Fix: Return an empty list if values contain None

for numeroPedido, grouped_row in grouped_rows.items():
    pedidos = []
    for row in grouped_row:
        numeroPedido, cnpjFornecedor, vendedor, valorFrete, valorImpostos, tipoPedido, tipoPrazoPagamento, prazosPagamento, codigoProduto, quantidade, precoBruto, descontoTotal,  aliquotaICMS, aliquotaIPI, previsaoEntrega, concluirPedido = row
        pedidos.append({
            'numeroPedido': numeroPedido,
            'cnpjFornecedor': cnpjFornecedor,
            'vendedor': vendedor,
            'valorFrete': valorFrete,
            'valorImpostos': valorImpostos,
            'tipoPedido': tipoPedido,
            'tipoPrazoPagamento': tipoPrazoPagamento,
            'prazosPagamento': prazosPagamento,
            'codigoProduto': codigoProduto,
            'quantidade': quantidade,
            'precoBruto': precoBruto,
            'descontoTotal': descontoTotal,
            'aliquotaICMS': aliquotaICMS,
            'aliquotaIPI': aliquotaIPI,
            'previsaoEntrega': previsaoEntrega,
            'concluirPedido': concluirPedido
        })
    
    results.append(pedidos)
    
    url = "http://ws.kplcloud.onclick.com.br/AbacosWSerp.asmx"
    headers = {'content-type': 'text/xml'}
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

    response = requests.post(url=url, headers=headers, data=body)
    response_content = response.content.decode('utf-8')
    xml_io = StringIO(response_content)
    for event, element in ET.iterparse(xml_io, events=("start", "end")):
        if event == "start" and element.tag == '{http://www.kplsolucoes.com.br/ABACOSWebService}InserirPedidoCompraResult':
    
            codigo = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Codigo")
            descricao = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Descricao")
            tipo = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Tipo")
            exceptionMessage = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}ExceptionMessage")
            # if codigo == "100001":
            print(codigo, descricao, tipo, exceptionMessage)
           
            
    for pedido in pedidos:
        
        handle_soap_request(url, headers, body)
        
