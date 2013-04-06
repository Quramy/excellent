package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jggug.kobo.gexcelapi.GExcel;

class Excellent {

	static{
		Cell.mixin(CellCategory)
		Row.mixin(RowCategory) 
		Sheet.mixin(SheetCategory)
	}

	static open(String file){
		GExcel.open(file)
	}
	static open(File file){
		GExcel.open(file)
	}
	static open(InputStream is){
		GExcel.open(is)
	}
}
