package net.quramy.excellent

import net.quramy.excellent.Excellent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;


class RowCategoryTest extends GroovyTestCase{	
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
