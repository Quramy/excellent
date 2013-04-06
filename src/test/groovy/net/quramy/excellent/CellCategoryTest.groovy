package net.quramy.excellent

import net.quramy.excellent.Excellent;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;


class CellCategoryTest extends GroovyTestCase{

	@Test
	void testLeft01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx")["Transition"]
		assertEquals("left", sheet.B2.left.value)
	}
	@Test
	void testLeft02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx")["Transition"]
		assertEquals("left", sheet.A2.left.value)
	}
	@Test
	void testRight01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx")["Transition"]
		assertEquals("right", sheet.B2.right.value)
	}
	@Test
	void testUp01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx")["Transition"]
		assertEquals("up", sheet.B2.up.value)
	}
	@Test
	void testUp02(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx")["Transition"]
		assertEquals("up", sheet.B1.up.value)
	}
	@Test
	void testDown01(){
		Sheet sheet = Excellent.open("src/test/resources/test.xlsx")["Transition"]
		assertEquals("down", sheet.B2.down.value)
	}
}
