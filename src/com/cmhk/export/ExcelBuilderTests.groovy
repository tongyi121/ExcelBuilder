package com.cmhk.export

import java.io.File;

import groovy.util.GroovyTestCase

class ExcelBuilderTests extends GroovyTestCase {

	File excel
	
	void setUp() {
		excel = new File("部门考评反馈表.xls")
		new File('out.xls').delete()
	}

	void testCreateSimpleWorkbook()  {
		def workbook = new ExcelBuilder(excel).workbook{
			sheet("部門匯總"){
				row(1){
					cell('A'){
						template([orgName:'信息技术部',year:2012])
					}
				}
				iRow(3){
					cell 'A','张三'
					cell 'B',90
					cell 'C',1
					cell 'D',90
				}
				
				row(4){
					cell 'A','李四'
					cell('B',90)
					cell('C',2)
					cell('D',90)
				}
			}
			
			cSheet('name','李四'){
				row(1){
					cell('B','李四')
					cell('F','信息技术部')
				}
				row(2){
					cell('B','普通职员')
					cell('F','2012年度')
				}
			}
			
			sheet('name','张三'){
				row(1){
					cell('B','张三')
					cell('F','信息技术部')
				}
				row(2){
					cell('B','普通职员')
					cell('F','2012年度')
				}
			}
			
			
		}
		new File('out.xls').withOutputStream {
			workbook.write(it)
		}
		
		
	}
}
