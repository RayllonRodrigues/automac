import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === Função de limpeza de texto ===
def limpar(texto):
    return re.sub(r'[\xa0\n\r]+', ' ', texto).strip()

# === DADOS DE LOGIN ===
email = "email"
senha = "suasenha"

# === Iniciar navegador ===
driver = webdriver.Chrome()
driver.get("https://app.beesweb.com.br/")
time.sleep(3)

# === Login ===
driver.find_element(By.ID, "email").send_keys(email)
driver.find_element(By.ID, "password").send_keys(senha)
driver.find_element(By.ID, "password").submit()
time.sleep(5)

# === Abrir menu lateral ===
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "a.navbar-brand.d-flex"))
).click()
time.sleep(1)

# === Acessar aba "Clientes" ===
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "nav-link") and contains(text(), "Clientes")]'))
).click()
time.sleep(3)

# === PAGAMENTOS HOJE ===
div_total = driver.find_element(By.CLASS_NAME, "card-footer")
texto_pagamento = limpar(div_total.text)

# === STATUS DOS CLIENTES (gráfico de pizza) ===
total_geral = limpar(driver.find_element(By.CLASS_NAME, "total-value").text)
legenda = driver.find_elements(By.CLASS_NAME, "legend-item")

clientes_status = [{
    "Status": "TOTAL CLIENTES",
    "Quantidade": total_geral,
    "Percentual": "100%"
}]

for item in legenda:
    valor = limpar(item.find_element(By.CLASS_NAME, "item-value").text)
    label = limpar(item.find_element(By.CLASS_NAME, "item-label").text)
    percentual = limpar(item.find_element(By.CLASS_NAME, "item-percent").text)
    clientes_status.append({
        "Status": label,
        "Quantidade": valor,
        "Percentual": percentual
    })

# === ATIVAÇÕES DO MÊS ===
qtd_ativacoes = limpar(driver.find_element(
    By.XPATH, '//span[contains(text(), "Ativações do mês/ano")]/preceding-sibling::h3').text)

itens_ativacao = driver.find_elements(
    By.XPATH, '//span[contains(text(), "Ativações do mês/ano")]/ancestor::div[contains(@class, "card")]/div[contains(@class, "card-body")]//li')

ativacoes = [{"Nome": "Total Ativações", "Data de Cadastro": qtd_ativacoes}]
for item in itens_ativacao:
    texto = limpar(item.text)
    nome, data = texto.rsplit(" ", 1)
    ativacoes.append({"Nome": nome.strip(), "Data de Cadastro": data.strip()})

# === DESATIVAÇÕES DO MÊS ===
qtd_desativacoes = limpar(driver.find_element(
    By.XPATH, '//span[contains(text(), "Desativações do mês/ano")]/preceding-sibling::h3').text)

itens_desativacao = driver.find_elements(
    By.XPATH, '//span[contains(text(), "Desativações do mês/ano")]/ancestor::div[contains(@class, "card")]/div[contains(@class, "card-body")]//li')

desativacoes = [{"Nome": "Total Desativações", "Data de Desativação": qtd_desativacoes}]
for item in itens_desativacao:
    texto = limpar(item.text)
    nome, data = texto.rsplit(" ", 1)
    desativacoes.append({"Nome": nome.strip(), "Data de Desativação": data.strip()})

# === CONTADORES DO DIA ===
data_hoje = pd.Timestamp.now().strftime("%d/%m")
ativacoes_hoje = [c for c in ativacoes[1:] if c["Data de Cadastro"] == data_hoje]
desativacoes_hoje = [c for c in desativacoes[1:] if c["Data de Desativação"] == data_hoje]

# === SALVAR EM EXCEL ===
with pd.ExcelWriter("dashboard_beesweb.xlsx", engine="openpyxl") as writer:
    pd.DataFrame([{"Data": pd.Timestamp.now().strftime("%Y-%m-%d"), "Resumo": texto_pagamento}])\
        .to_excel(writer, sheet_name="pagamentos_hoje", index=False)

    pd.DataFrame(clientes_status).to_excel(writer, sheet_name="clientes_status", index=False)
    pd.DataFrame(ativacoes).to_excel(writer, sheet_name="ativacoes_mes", index=False)
    pd.DataFrame(desativacoes).to_excel(writer, sheet_name="desativacoes_mes", index=False)

    pd.DataFrame([{
        "Data": pd.Timestamp.now().strftime("%Y-%m-%d"),
        "Ativações Hoje": len(ativacoes_hoje),
        "Desativações Hoje": len(desativacoes_hoje)
    }]).to_excel(writer, sheet_name="resumo_hoje", index=False)

print("\u2705 Arquivo 'dashboard_beesweb.xlsx' gerado com 5 abas.")
driver.quit()
