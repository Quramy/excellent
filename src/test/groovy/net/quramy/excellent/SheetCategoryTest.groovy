package net.quramy.excellent

import net.quramy.excellent.Excellent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;


class SheetCategoryTest extends GroovyTestCase{

	@Test
	void testFind01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell cell = sheet.find{it=~/edge/}
		assertEquals("A1", cell.label)
	}
	
	@Test
	void testFind02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell cell = sheet.find{it=~/nobody/}
		assertNull(cell)
	}
	
	@Test
	void testTop01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		assertTrue sheet.getRow(0) == sheet.top
	}
	
	@Test
	void testBottom01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		assertTrue sheet.getRow(5) == sheet.bottom
	}

}
