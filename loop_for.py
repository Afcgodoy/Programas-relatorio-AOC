#Para escrever de 0 a 10000 em um arquivo txt:

def escrever_arquivo(numero, arquivo = 'numeros.txt'):
    with open(arquivo, 'a') as arq:
        arq.write(f"{numero}\n")
        
for i in range(0, 10001):
    escrever_arquivo(i)       