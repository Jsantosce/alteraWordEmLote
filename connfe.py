import tkinter as tk
import tkinter.filedialog
import os
import requests
from xml.etree import ElementTree as ET
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
from OpenSSL.crypto import load_pkcs12, FILETYPE_PEM

# Função para abrir a janela de seleção de arquivo do certificado
def browse_certificate():
    filetypes = (("PKCS12 Files", "*.p12"), ("PEM Files", "*.pem"), ("All Files", "*.*"))
    certificate_file = tkinter.filedialog.askopenfilename(title="Select Certificate File", filetypes=filetypes)
    if certificate_file:
        certificate_file_entry.delete(0, tk.END)
        certificate_file_entry.insert(0, certificate_file)

# Função para consultar a NF-e
def consult_nfe():
    nfe_key = nfe_key_entry.get()
    certificate_file = certificate_file_entry.get()
    certificate_password = certificate_password_entry.get()

    # Verifica se os campos foram preenchidos
    if not nfe_key or not certificate_file or not certificate_password:
        tk.messagebox.showerror("Error", "Please fill all fields!")
        return

    # Carrega o arquivo do certificado
    try:
        if certificate_file.lower().endswith(".pem"):
            with open(certificate_file, "rb") as f:
                certificate = load_pem_x509_certificate(f.read(), default_backend())
                private_key = None
        else:
            with open(certificate_file, "rb") as f:
                p12 = load_pkcs12(f.read(), certificate_password.encode())
                certificate = p12.get_certificate()
                private_key = p12.get_privatekey()
    except Exception as e:
        tk.messagebox.showerror("Error", "Failed to load certificate: " + str(e))
        return

    # Realiza a consulta da NF-e
try:
    url = "https://nfe.sefaz.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
    headers = {"Content-Type": "application/soap+xml;charset=utf-8"}
    body = f"""<?xml version="1.0" encoding="UTF-8"?>
            <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:nfe="http://www.portalfiscal.inf.br/nfe/wsdl/NfeConsulta4">
                <soap:Header>
                    <nfe:NfeCabecMsg xmlns:nfe="http://www.portalfiscal.inf.br/nfe/wsdl/NfeConsulta4">
                        <nfe:versaoDados>4.00</nfe:versaoDados>
                        <nfe:certificado>{certificate.public_bytes().hex()}</nfe:certificado>
                    </nfe:NfeCabecMsg>
                </soap:Header>
                <soap:Body>
                    <nfe:consSitNFe xmlns:nfe="http://www.portalfiscal.inf.br/nfe" versao="4.00">
                        <nfe:tpAmb>1</nfe:tpAmb>
                        <nfe:xServ>CONSULTAR</nfe:xServ>
                        <nfe:chNFe>{nfe_key}</nfe:chNFe>
                    </nfe:consSitNFe>
                </soap:Body>
            </soap:Envelope>"""

    response = requests.post(url, headers=headers, data=body)

    if response.status_code == 200:
        # Parseia a resposta
        response_body = response.content.decode("utf-8")
        xml = ET.fromstring(response_body)

        if xml.find(".//cStat").text == "100":
            # Se a consulta for bem-sucedida, salva o XML em disco
            filetypes = (("XML Files", "*.xml"), ("All Files", "*.*"))
            filename = tkinter.filedialog.asksaveasfilename(title="Save XML", filetypes=filetypes, defaultextension=".xml")
            if filename:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(response_body)
                tk.messagebox.showinfo("Success", "XML saved successfully!")
            else:
                # Se a consulta falhar, exibe a mensagem de erro
                tk.messagebox.showerror("Error", xml.find(".//xMotivo").text)
            except Exception as e:
            # Exibe a mensagem de erro genérica
            tk.messagebox.showerror("Error", f"Failed to consult NF-e: {str(e)}")
            response = requests.post(url, headers=headers, data=body.encode(), cert=(certificate, private_key))

            # Verifica se a resposta da SEFAZ foi bem sucedida
            if response.status_code == 200:
                response_tree = ET.fromstring(response.content)
                nfe_consulta_result = response_tree.find(".//{http://www.portalfiscal.inf.br/nfe}nfeConsultaNFResult")
                if nfe_consulta_result is not None:
                    xml_string = nfe_consulta_result.text.strip()
                    # Salva o XML da NF-e em um arquivo temporário
                    with open("nfe_temp.xml", "w") as f:
                        f.write(xml_string)
                    tk.messagebox.showinfo("Success", "NF-e found! Check the downloaded XML.")
                else:
                    tk.messagebox.showinfo("Not Found", "NF-e not found.")
            else:
                tk.messagebox.showerror("Error", f"Failed to query NF-e. Error code: {response.status_code}")
            except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to query NF-e: {str(e)}")
            return

        # Trata erros de validação do XML
            except ET.ParseError as e:
            tk.messagebox.showerror("Error", "Failed to parse response from SEFAZ: " + str(e))
            return


root = tk.Tk()
root.title("NF-e Query")

nfe_key_label = tk.Label(root, text="NF-e Key:")
nfe_key_entry = tk.Entry(root, width=40)
certificate_file_label = tk.Label(root, text="Certificate File:")
certificate_file_entry = tk.Entry(root, width=40)
certificate_file_button = tk.Button(root, text="Browse", command=browse_certificate)
certificate_password_label = tk.Label(root, text="Certificate Password:")
certificate_password_entry = tk.Entry(root, width=40, show="*")
consult_button = tk.Button(root, text="Consult NF-e", command=consult_nfe)
quit_button = tk.Button(root, text="Quit", command=root.quit)

nfe_key_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
nfe_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
certificate_file_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
certificate_file_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
certificate_file_button.grid(row=1, column=2, padx=5, pady=5)
certificate_password_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
certificate_password_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
consult_button.grid(row=3, column=1, padx=5, pady=5)
quit_button.grid(row=3, column=2, padx=5, pady=5)

root.mainloop()
