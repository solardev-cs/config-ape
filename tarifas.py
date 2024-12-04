import requests

# VARIÁVEIS API
URL1 = "https://apise.way2.com.br/v1/agentes"
URL2 = "https://apise.way2.com.br/v1/tarifas"
API = "a83ecad28f3b41f88590db10a9b675f4"

# OBTÉM LISTA DITRIBUIDORAS DA API
def get_distrib():
    concessionarias = []

    parametros = {"apikey": API}
    response = requests.get(URL1, params=parametros, timeout=3)
    data = response.json()

    for item in data:
        concessionarias.append(item["nome"])
    
    return concessionarias

# OBTÉM ANO VÁLIDO MAIS RECENTE DA API
def get_ano_tarifas(distrib, ano_planilha):
    ano = ano_planilha
    data = None

    while not data:
        parametros = {"apikey": API,
                    "agente": distrib,
                    "ano": ano}
        response = requests.get(URL2, params=parametros, timeout=3)
        data = response.json()
        ano -= 1

    return (ano+1)

# OBTÉM TARIFAS DA API
def get_tarifas(distrib, ano, subgrupo, mod):
    tarifas = {}
    
    parametros = {"apikey": API,
                  "agente": distrib,
                  "ano": ano}
    response = requests.get(URL2, params=parametros, timeout=3)
    data = response.json()

    if mod == "VERDE":
        mod_ape = "VERDE APE"
    elif mod == "AZUL":
        mod_ape = "AZUL APE"

    for item in data:
        
        if item["subgrupo"] == subgrupo:   

            if item["modalidade"].upper() == mod:
                # salva tarifas verde ou azul normal - consumo HFP, HP e demanda TUSD D
                if "VERDE" in item["modalidade"].upper():
                    if item["posto"] == "FP":
                        tarifas["TUSD_C_HFP"] = item["tarifaconsumotusd"]
                    elif item["posto"] == "P":
                        tarifas["TUSD_C_HP"] = item["tarifaconsumotusd"]
                    elif item["posto"] == "NA":
                        tarifas["TUSD_D"] = item["tarifademandatusd"]
                elif "AZUL" in item["modalidade"].upper():
                    if item["posto"] == "FP":
                        tarifas["TUSD_C_HFP"] = item["tarifaconsumotusd"]
                        tarifas["TUSD_D"] = item["tarifademandatusd"]
                    elif item["posto"] == "P":
                        tarifas["TUSD_C_HP"] = item["tarifaconsumotusd"]

            elif item["modalidade"].upper() == mod_ape:
                # salva tarifas verde ou azul APE - consumo HFP e HP
                if item["posto"] == "FP":
                    tarifas["TUSD_APE_HFP"] = item["tarifaconsumotusd"]
                elif item["posto"] == "P":
                    tarifas["TUSD_APE_HP"] = item["tarifaconsumotusd"]

            elif item["modalidade"].upper() == "GERAÇÃO":
                # salva tarifa demanda TUSD G
                tarifas["TUSD_G"] = item["tarifademandatusd"]                

    if not tarifas:
        return {"erro": f"Subgrupo '{subgrupo}' não existente nessa distribuidora. Selecione um subgrupo válido e clique em Salvar novamente."}

    return tarifas