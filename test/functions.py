# -*- coding: utf-8 -*-

def F(x):
	y = zeros(4, dtype=x)
	y[0] = x[0] - (x[1] + x[2])
	y[1] = x[1] - x[3]
	y[2] = x[2] - x[4]
	y[3] = x[3] + x[4] - x[5]
	return y