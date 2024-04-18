import logging
import requests
from io import StringIO
import pandas as pd
from typing import List, Tuple
import xml.etree.ElementTree as ET

def setSentStatus() -> List[Tuple[str]]:
    try:
        workbook = pd.read_excel("C:/orders/Pasta1b.xlsx")
        rows = []
        for index, row in workbook.iterrows():
            numeroPedido = str(row.iloc[0])
            
            column1 = row.iloc[1]
            column2 = row.iloc[2]
            column3 = row.iloc[3]
            
            rows.append((numeroPedido, column1, column2, column3))
        return rows
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        return []
rows = setSentStatus()
results = []

for row in rows:
    numeroPedido = row
    url = "http://ws.kplcloud.onclick.com.br/AbacosWswms.asmx"
    headers = {'content-type': 'text/xml'}
    body = f"""<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
            <InserirPedidoCompra xmlns="http://www.kplsolucoes.com.br/ABACOSWebService">
                <ChaveIdentificacao>E19C1DD8-345F-4F1B-A53F-CA4312CAF457</ChaveIdentificacao>
                    <ListaDePedidosCompra>
                        <PedidoCompra>
                            <NumeroPedido>string</NumeroPedido>
                            <CnpjFornecedor>string</CnpjFornecedor>
                            <Vendedor>string</Vendedor>
                            <ValorFrete>string</ValorFrete>
                            <ValorImpostos>string</ValorImpostos>
                            <TipoPedido>string</TipoPedido>
                            <TipoPrazoPagamento>string</TipoPrazoPagamento>
                            <PrazosPagamento>string</PrazosPagamento>
                            <CodigoProduto>string</CodigoProduto>
                            <Quantidade>string</Quantidade>
                            <PrecoBruto>string</PrecoBruto>
                            <DescontoTotal>string</DescontoTotal>
                            <AliquotaICMS>string</AliquotaICMS>
                            <AliquotaIPI>string</AliquotaIPI>
                            <PrevisaoEntrega>string</PrevisaoEntrega>
                            <ConcluirPedido>string</ConcluirPedido>
                        </PedidoCompra>
                    </ListaDePedidosCompra>
            </InserirPedidoCompra>
        </soap:Body>
    </soap:Envelope>"""

    def handle_soap_request(url: str, headers: dict, body: str) -> List[Tuple[str]]:
        try:
            response = requests.post(url, data=body, headers=headers)
            response_content = response.content.decode('utf-8')
            response_content = ''.join(char for char in response_content if ord(char) < 255)
            # print(response_content)

            xml_io = StringIO(response_content)
            result = []
            values = []

            for event, element in ET.iterparse(xml_io, events=("start", "end")):
                if event == "start" and element.tag == '{http://www.kplsolucoes.com.br/ABACOSWebService}MarcarPedidosDespachados':
                    values = tuple(
                        element.findtext(f".//{{http://www.kplsolucoes.com.br/ABACOSWebService}}{tag}")
                        for tag in [
                            "NumeroDoPedido"
                        ]
                    )
                    # for event, element in ET.iterparse(xml_io, events=("start", "end")):
                    #     if event == "start" and element.tag == '{http://www.kplsolucoes.com.br/ABACOSWebService}MarcarPedidosDespachadosResult':
                    #         codigo = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Codigo")
                    #         descricao = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Descricao")
                    #         tipo = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}Tipo")
                    #         exceptionMessage = element.findtext(".//{http://www.kplsolucoes.com.br/ABACOSWebService}ExceptionMessage")
                    #         if codigo == "200001":
                    #             print(codigo, descricao, tipo, exceptionMessage)
                    #         else:
                    #             logging.warning("Error processing row: %s", exceptionMessage)
                            
            if None not in values:
                return result
            result.append(values)
            
                    
        except requests.exceptions.RequestException as e:
            print(f"An error occurred: {e}")
        
        return []  # Fix: Return an empty list if values contain None
    
    result = handle_soap_request(url, headers, body)
    results.extend(result)
