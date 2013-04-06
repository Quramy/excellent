package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

@Category(Sheet)
class SheetCategory {
	
	Object find(Closure<?> closure){
		Cell res = null
		int s = this.firstRowNum
		int e = this.lastRowNum
		for(int i = s; i <= e; i++){
			Row row = this.getRow i
			Cell tmp = row?.find(closure)
			if(tmp != null){
				res = tmp
				break
			}
		}
		return res
	}
	
	Row getTop(){
		return this?.getRow(this.firstRowNum)
	}
	
	Row getBottom(){
		return this?.getRow(this.lastRowNum)
	}
}
