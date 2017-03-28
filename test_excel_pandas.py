#-*-coding:utf-8-*-

from excel_pandas import *
import unittest
import os
import pandas as pd
import numpy as np

class BasicTestCase(unittest.TestCase):
	def setUp(self):
		#建立两个EXCEL文件
		df1 = pd.DataFrame({
			'Date':pd.date_range('20170102',periods=10),
			'A':range(10),
			'B':['a'*5,'n'*5]*5,
			'Data1':[1,2,3,4,5]*2,
			'Data2':[1,3,5,7,9]*2,
			})
		df2 = pd.DataFrame({
			'Date':pd.date_range('20170102',periods=10),
			'A':range(20,30),
			'B':['a','b','c','d','e']*2,
			'Data1':[5,4,3,2,1]*2,
			'Data2':[1,2,3,4,5,6,7,8,9,10],
			})
		df1 = df1.set_index(['Date'])
		df2 = df2.set_index(['Date'])
		df1.to_excel('t1.xlsx')
		df2.to_excel('t2.xlsx')
		
	def tearDown(self):
		os.remove('t1.xlsx')
		os.remove('t2.xlsx')
		os.remove('new.xlsx')
		
	def test_set_path(self):
		a1 = MakePivotByPandas()
		a1.run()
		self.assertEquals(a1.path,basedir)
		a2 = MakePivotByPandas('\Python27\project')
		self.assertEquals(a2.path,'\Python27\project')
		
	def test_read_data(self):
		a = MakePivotByPandas()
		a.run()
		self.assertEquals(len(a.datas),2)
		#print a.datas
		self.assertEquals(a.datas[0]['Data1'][0],a.datas[1]['Data2'][0])
		self.assertTrue(a.datas[0].columns[0] == a.datas[1].columns[0])
		
	def test_make_pivot_table(self):
		a = MakePivotByPandas()
		a.run()
		self.assertEquals(len(a.tables),2)
		print a.tables
		self.assertTrue(a.tables[0].columns[0] == 'Data1')
		self.assertEquals(a.tables[0]['Data1'][0],1)
		self.assertTrue(a.tables[1]['Data1'][9],1)
		
	def test_merge_table(self):
		a = MakePivotByPandas()
		a.run()
		#print a.new_table.index
		self.assertTrue(a.new_table['Data1'][0])
		self.assertTrue(a.new_table.index[0]==pd.Timestamp('20170102'))
		
		self.assertEquals(a.new_table['Data1'][0],6)
		
	def test_save(self):
		a = MakePivotByPandas()
		a.run()
		self.assertTrue('new.xlsx' in os.listdir(basedir))
		
if __name__ == '__main__':
	unittest.main()
		
