# -*- coding: utf-8 -*-
from cvxopt import matrix, solvers, spmatrix
from cvxopt.lapack import getrf, getri
from algopy import UTPM


class pyRecon:

	def __init__(self, x, QI, lb, ub, F):
		self.F = F
		self.x = matrix(x)
		self.QI = matrix(QI)
		self.lb = matrix(lb)
		self.ub = matrix(ub)
		self.N = self.x.size[0]
		self.P = spmatrix(1.0, list(range(self.N)), list(range(self.N)), tc='d')
		self.G = matrix([-self.P, self.P])
		self.h = matrix([self.ub - self.x, self.x - self.lb])
		self.A = self.obterIncidencia()
		self.y = None
		self.QIr = None

	def calcularQIReconciliada(self):
		" Calcula a QI após a reconciliação de dados"
		I = spmatrix(1.0, range(self.N), range(self.N), tc='d')
		U = matrix(0.0, (self.N, self.N), 'd')
		self.QIr = matrix(0.0, (self.N + 1, 1), 'd')

		for i in range(self.N):
			U[i, i] = (0.1 * self.x[i] / (self.QI[i] * 0.9 * 3.0 ** 0.5)) ** 2

		ipiv = matrix(0, (self.N, 1))
		AI = self.A * U * self.A.T
		getrf(AI, ipiv)
		getri(AI, ipiv)

		S = I - U * self.A.T * AI * self.A
		Ur = S * U * S.T

		for i in range(self.N):
			self.QIr[i] = 0.1 * self.y[i] / (0.9 * (3.0 * Ur[i, i]) ** 0.5)

		self.QIr[self.N] = self.QIr[:self.N].T * self.y / sum(self.y)
		return

	def calcularReconciliacao(self, options):
		" Efetua a reconciliação de dados com QI"
		# Constrói opções do CVXOPT
		if options is None:  # Opções padrão
			solvers.options['show_progress'] = False
			solvers.options['reltol'] = 1e-10
			solvers.options['maxiters'] = 10000
		else:  # Opções foram passadas
			for i in range(0, len(options), 2):
				solvers.options[options[i]] = eval(options[i + 1])

		# Substitui a matriz identidade P pela matriz de incertezas
		for i in range(self.N):
			self.P[i, i] = (self.QI[i] / self.x[i]) ** 2

		# Adequando ao modelo padrão da QP
		# 1/2 x.T P x + q.T x
		# s.t.
		# G x <= h
		# A x = b
		q = matrix(0.0, (self.N, 1), 'd')
		self.P *= 2.0
		b = self.A * self.x

		# Resolvendo o sistema via CVXOPT
		sol = solvers.coneqp(P=self.P, q=q, A=self.A, b=b, G=self.G, h=self.h)
		
		# Corrigindo a solução ==> y_opt = x - y
		self.y = self.x - sol['x']
		return

	def obterIncidencia(self):
		"Obtém a matriz de incidência através das funções geradas"
		x0 = UTPM.init_jacobian(self.x)
		y0 = self.F(x0)
		return matrix(UTPM.extract_jacobian(y0))

	def resolver(self, options):
		self.calcularReconciliacao(options)
		self.calcularQIReconciliada()
		return