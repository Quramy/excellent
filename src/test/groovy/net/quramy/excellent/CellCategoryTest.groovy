package net.quramy.excellent

import net.quramy.excellent.Excellent;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Before;
import org.junit.Test;


class CellCategoryTest extends GroovyTestCase{
	
	def Sheet sheet 

	@Before
	@Override
	void setUp(){
		sheet = Excellent.open("src/test/resources/test.xlsx").Transition
	}
	
	@Test
	void testLeft01(){
		assertEquals("left", sheet.B2.left.value)
	}
	@Test
	void testLeft02(){
		assertEquals("A5", sheet.C5.left.left.label)
	}
	@Test
	void testRight01(){
		assertEquals("right", sheet.B2.right.value)
	}
	@Test
	void testRight02(){
		assertEquals("C5", sheet.A5.right.right.label)
	}
	@Test
	void testUp01(){
		assertEquals("up", sheet.B2.up.value)
	}
	@Test
	void testUp02(){
		assertEquals("A2", sheet.A4.up.up.label)
	}
	@Test
	void testDown01(){
		assertEquals("down", sheet.B2.down.value)
	}
	@Test
	void testDown02(){
		assertEquals("A4", sheet.A2.down.down.label)
	}
	
	@Test
	void testPlus01(){
		assertEquals("B1", (sheet.A1 + 1).label);
	}
	@Test
	void testMinus01(){
		assertEquals("A1", (sheet.B1 - 1).label);
	}
	@Test
	void testMinus02(){
		assertEquals("A1", (sheet.A1 - 1).label);
	}
}
