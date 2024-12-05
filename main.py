import json
import pandas as pd
from colorama import init, Fore, Style


init(autoreset=True)


def salvar_dados(alunos):
    try:
        with open("dados_alunos.json", "w") as f:
            json.dump(alunos, f, indent=4)
        print(f"{Fore.GREEN}Dados salvos com sucesso!{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}Erro ao salvar os dados: {e}{Style.RESET_ALL}")


def carregar_dados():
    try:
        with open("dados_alunos.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return []  
    except json.JSONDecodeError:
        print(f"{Fore.RED}Erro ao ler o arquivo JSON. O arquivo pode estar corrompido.{Style.RESET_ALL}")
        return [] 


def exportar_para_excel(alunos):
    if not alunos:
        print(f"{Fore.RED}Nenhum dado de aluno disponível para exportar.{Style.RESET_ALL}")
        return

    
    for aluno in alunos:
        aluno["notas"] = ', '.join(map(str, aluno["notas"]))  

    try:
       
        df = pd.DataFrame(alunos)
        df.to_excel("relatorio_alunos.xlsx", index=False, engine='openpyxl')
        print(f"{Fore.GREEN}Relatório exportado com sucesso para 'relatorio_alunos.xlsx'.{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}Erro ao exportar para Excel: {e}{Style.RESET_ALL}")


def main():
    alunos = carregar_dados()

    while True:
        print("\nMenu Principal:")
        print("1. Cadastrar novo aluno")
        print("2. Exportar relatório para Excel")
        print("3. Sair")
        opcao = input("Escolha uma opção: ").strip()

        if opcao == "1":
            nome = input("Digite o nome do aluno: ").strip()
            notas = []
            for i in range(1, 5):
                while True:
                    try:
                        nota = float(input(f"Digite a nota do {i}º bimestre (0 a 10): "))
                        if 0 <= nota <= 10:
                            notas.append(nota)
                            break
                        else:
                            print(f"{Fore.RED}A nota deve estar entre 0 e 10.{Style.RESET_ALL}")
                    except ValueError:
                        print(f"{Fore.RED}Por favor, insira um número válido para a nota.{Style.RESET_ALL}")
            media = sum(notas) / len(notas)

            aluno = {"nome": nome, "notas": notas, "media": round(media, 2)}
            alunos.append(aluno)
            salvar_dados(alunos)
            print(f"{Fore.GREEN}Aluno cadastrado com sucesso!{Style.RESET_ALL}")

        elif opcao == "2":
            exportar_para_excel(alunos)

        elif opcao == "3":
            print(f"{Fore.CYAN}Encerrando o programa. Até mais!{Style.RESET_ALL}")
            break

        else:
            print(f"{Fore.RED}Opção inválida! Tente novamente.{Style.RESET_ALL}")


if __name__ == "__main__":
    main()