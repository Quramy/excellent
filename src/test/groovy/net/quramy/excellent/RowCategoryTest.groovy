package net.quramy.excellent

import net.quramy.excellent.Excellent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;


class RowCategoryTest extends GroovyTestCase{

	@Test
	void testFind01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell res = sheet.getRow(1).find{it=~/center/}
		assertEquals("B2", res.label)
	}
	
	@Test
	void testFind02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell res = sheet.getRow(1).find{it=~/nobody/}
		assertNull(res)
	}
	
	@Test
	void testFind03(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell res = sheet.getRow(5).find{it =~/edge/}
		assertEquals("D6", res.label)
	}
	
	
	@Test
	void testFindAll01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		def res = sheet.getRow(4).findAll{it =~/hoge/}
		assertEquals(["A5", "C5"], res*.label)
	}
	
	
	@Test
	void testFindAll02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		def res = sheet.getRow(3).findAll{it =~/hoge/}
		assertEquals([], res*.label)
	}
	
	@Test
	void testFirst01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		assertEquals("A2", sheet.getRow(1).first.label)
	}
	
	@Test
	void testLast01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		assertEquals("C2", sheet.getRow(1).last.label)
	}
}
