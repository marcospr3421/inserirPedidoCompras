from fileinput import filename
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



def gera_pedido_compra(pedidos: List[dict]) -> str:
    pedidos_compra = ""
    
    for pedido in pedidos:

        pedidos_compra += f"""
        <PedidoCompra>
            <NumeroPedido>{pedido['numeroPedido']}</NumeroPedido>
            <CnpjFornecedor>{pedido['cnpjFornecedor']}</CnpjFornecedor>
            <Vendedor>{pedido['vendedor']}</Vendedor>
            <ValorFrete>{pedido['valorFrete']}</ValorFrete>
            <ValorImpostos>{pedido['valorImpostos']}</ValorImpostos>
            <TipoPedido>{pedido['tipoPedido']}</TipoPedido>
            <TipoPrazoPagamento>{pedido['tipoPrazoPagamento']}</TipoPrazoPagamento>
            <PrazosPagamento>{pedido['prazosPagamento']}</PrazosPagamento>
            <CodigoProduto>{pedido['codigoProduto']}</CodigoProduto>
            <Quantidade>{pedido['quantidade']}</Quantidade>
            <PrecoBruto>{pedido['precoBruto']}</PrecoBruto>
            <DescontoTotal>{pedido['descontoTotal']}</DescontoTotal>
            <AliquotaICMS>{pedido['aliquotaICMS']}</AliquotaICMS>
            <AliquotaIPI>{pedido['aliquotaIPI']}</AliquotaIPI>
            <PrevisaoEntrega>{pedido['previsaoEntrega']}</PrevisaoEntrega>
            <ConcluirPedido>{pedido['concluirPedido']}</ConcluirPedido>
        </PedidoCompra>"""
    return pedidos_compra
    


    

def read_excel_file(filename):
    try:
        workbook = pd.read_excel(filename)
        rows = []
        for index, row in workbook.iterrows():
            rows.append(tuple(str(row.iloc[i]) for i in range(16)))
        return rows
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        return []
def select_file():
    global filename
    filename = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel Files", "*.xlsx")])

window = tk.Tk()
window.title("File Selector")

filename = ''  # Initialize filename as an empty string
button1 = tk.Button(window, text="Select Excel File", command=select_file)
button1.pack(side=tk.LEFT)

button2 = tk.Button(window, text="Import", command=lambda: read_excel_file(filename))
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
        
