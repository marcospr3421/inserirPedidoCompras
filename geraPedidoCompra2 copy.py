import logging
import requests
from io import StringIO
import pandas as pd
from typing import List, Tuple
from collections import defaultdict
from tkinter import Tk, filedialog
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox
import xml.sax.saxutils as saxutils

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

def set_order(filename: str) -> List[Tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]]:
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
            rows.append(tuple(str(row.iloc[i]) for i in range(16)))
        return rows
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        return []


def handle_soap_request(url: str, headers: dict, body: str) -> List[Tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str]]:
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
                        "ConcluirPedido"
                    ]
                )

                if None not in values:
                    result.append(values)

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")

    return result

def main():
    """
    Main function of the script.
    """

    # Select the Excel file
    filename = select_file()
    if not filename:
        return

    # Set the purchase order data
    rows = set_order(filename)
    if not rows:
        return

    # Group rows by purchase order number
    grouped_rows = defaultdict(list)
    for row in rows:
        grouped_rows[row[0]].append(row)

    # Process each purchase order
    for numero_pedido, grouped_row in grouped_rows.items():
        pedidos = []
        for row in grouped_row:
            numero_pedido, cnpj_fornecedor, vendedor, valor_frete, valor_impostos, tipo_pedido, tipo_prazo_pagamento, prazos_pagamento, codigo_produto, quantidade, preco_bruto, desconto_total, aliquota_icms, aliquota_ipi, previsao_entrega, concluir_pedido = row
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
                'concluirPedido': concluir_pedido
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
        url = "http://ws.kplcloud.onclick.com.br/AbacosWSerp.asmx"
        headers = {'content-type': 'text/xml'}
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



if __name__ == "__main__":
    main()
