import os

# Execute o primeiro código
os.system('python AtualizarDadosDiarioOFICIAL.py')

# Execute o segundo código após o término do primeiro
os.system('python CalculoMaiorLatencia.py')

print("Ambos os códigos foram executados em sequência.")
