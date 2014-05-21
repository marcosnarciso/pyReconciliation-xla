# -*- coding: utf-8 -*-

#import os
from win32api import MessageBox
from algopy import zeros
from pyRecon import pyRecon
from re import sub


def calculateReconciliation(path, options):
	"""
	Efetua a reconciliação de dados utilizando o CVXOPT
	O sistema é representado pelos arquivos data.py e functions.py
	path: Endereço da pasta onde os arquivos data.py e functions.py se encontram
	options: Vetor de opções do solver
	"""
	# Opções
	# Cria o dicionário de opções do solver passados pelo Excel
	opcoes = {}
	for i in range(0, len(options), 2):
		opcoes[options[i]] = options[i + 1]

	# Se a chamada contiver uma fórmula CÉL, separa a pasta do nome do arquivo
	#folder, sep, plan = path.partition("[")
	#data = folder + os.path.sep + 'data.py'
	#functions = folder + os.path.sep + 'functions.py'
	# Chamada direta pelo VBA
	data = path + '\\data.py'
	functions = path + '\\functions.py'

	# Executa os arquivos em Python em bytecode
	exec(open(data).read())  # Data (x, QI, lb, ub)
	exec(open(functions).read())  # Funções (F[x])

	# Executa a reconciliação de dados via CVXOPT e calcula QI após a reconciliação
	rec = pyRecon.pyRecon(x, QI, lb, ub, F)
	rec.resolver(options)

	sol = list(rec.y)
	sol.extend(list(rec.QIr))
	return sol


def createFunctions(label, func, path):
	"""
	Cria o arquivo functions.py
	label: lista de variáveis
	func: lista de equações
	path: pasta onde será gerado o arquivo
	"""
	# Rearruma as funções oriundas do VBA
	func = appendFunctions(func)
	# Loop de substituição
	for i in range(len(label)):
		for j in range(len(func)):
			# Expressão regular: + TAG + <==> + x[i] +
			func[j][0] = sub(r'\b' + label[i][0] + r'\b', str('x[%d]' % i), func[j][0])

	# Se a chamada contiver uma fórmula CÉL, separa a pasta do nome do arquivo
	#folder, sep, plan = path.partition("[")
	#data = folder + os.path.sep + "functions.py"
	# Chamada direta por VBA
	data = path + "\\functions.py"

	# Abre o arquivo para escrita
	f = open(data, 'w')

	# Preenche y no formato do ALGOPY
	f.write("# -*- coding: utf-8 -*-" + "\n\n" + "def F(x):\n")
	f.write("\ty = zeros(" + str(len(func)) + ", dtype=x)\n")

	# Equações
	for i in range(len(func)):
		f.write(str('\ty[%d] = ' % i) + func[i][0] + "\n")

	f.write("\treturn y")

	f.close()

	return True


def createData(x, QI, lb, ub, path):
	"""
	Cria o arquivo data.py
	x: valores mapeados
	QI: valores de Qualidade da Informação
	lb: limites inferiores para x
	ub: limites superiores para x
	path: pasta onde o arquivo será gerado
	"""
	# Checa se os vetores admitidos possuem a mesma dimensão
	if (len(x) != len(QI)) or (len(x) != len(lb)) or (len(x) != len(ub)):
		MessageBox(0, u'Dimensões incompatíveis.', "Erro!")
		return False
	else:
		# Se a chamada contiver uma fórmula CÉL, separa a pasta do nome do arquivo
		#folder, sep, plan = path.partition("[")
		#data = folder + os.path.sep + "data.py"
		# Chamada direta por VBA
		data = path + "\\data.py"

		# Abre o arquivo para escrita
		f = open(data, 'w')

		# Valores de x
		# x = matrix([x[0], x[1], x[2]]).T
		f.write("# -*- coding: utf-8 -*-" + "\n\n" + "x = [")
		for i in range(len(x) - 1):
			f.write(str('%f,' % (0.000001 if abs(x[i][0]) <= 0.00001 else x[i][0])))
		f.write(str('%f]\n' % (0.000001 if abs(x[len(x) - 1][0]) <= 0.00001 else x[len(x) - 1][0])))

		# Valores de lb
		# lb = matrix([lb[0], lb[1], lb[2]]).T
		f.write("lb = [")
		for i in range(len(lb) - 1):
			f.write(str('%f,' % lb[i][0]))
		f.write(str('%f]\n' % lb[len(lb) - 1][0]))

		# Valores de ub
		# ub = matrix([ub[0], ub[1], ub[2]]).T
		f.write("ub = [")
		for i in range(len(ub) - 1):
			f.write(str('%f,' % ub[i][0]))
		f.write(str('%f]\n' % ub[len(ub) - 1][0]))

		# Valores de QI
		# QI = matrix([QI[0], QI[1], QI[2]]).T
		f.write("QI = [")
		for i in range(len(QI) - 1):
			f.write(str('%.6f,' % QI[i][0]))
		f.write(str('%.6f]\n' % QI[len(QI) - 1][0]))

		f.close()

		return True


# Agrupa as equações passadas pelo VBA
def appendFunctions(functions):
	j, aux, func = 0, "", []
	for i in range(len(functions)):
		# Situação neste trecho
		# função[0] ==> V1 + V2 @
		# funcao[1] ==> - V3@@END@@
		if functions[i].find("@@END@@") >= 0:
			#Passo 2
			# func[i] ==> V1 + V2 - V3@@END@@
			aux += functions[i][:functions[i].index("@@END@@")]
			# Passo 3
			# func[i] ==> V1 + V2 - V3
			func.append([aux])
			aux = ""
			j += 1
		else:
			#Passo 0
			# func[i] ==> V1 + V2 @
			aux += functions[i][:functions[i].index("@")]
			# Passo 1
			# func[i] ==> V1 + V2
	return func