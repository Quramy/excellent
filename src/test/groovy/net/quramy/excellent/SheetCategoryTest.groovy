package net.quramy.excellent

import net.quramy.excellent.Excellent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;


class SheetCategoryTest extends GroovyTestCase{

	@Test
	void testFindCell01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell cell = sheet.findCell{it=~/edge/}
		assertEquals("A1", cell.label)
	}
	
	@Test
	void testFindCell02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		Cell cell = sheet.find{it=~/nobody/}
		assertNull(cell)
	}
	
	@Test
	void testFindCellAll01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		def res = sheet.findCellAll{it=~/edge/}
		assertEquals(["A1", "D6"],  res*.label)
	}
	
	@Test
	void testFindCellAll02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Transition
		def res = sheet.findCellAll{it=~/nobody/}
		assertEquals([],  res*.label)
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
