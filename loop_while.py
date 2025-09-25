#Para escrever de 0 a 10000 em um arquivo txt:

def escrever_arquivo(numero, arquivo = 'numeros_2.txt'):
    with open(arquivo, 'a') as arq:
        arq.write(f"{numero}\n")
        
i = 0

while i<=10000:
    escrever_arquivo(i)
    i = i+1    