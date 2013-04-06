package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Override find, findAll method to Row.
 * @author yosuke
 *
 */
@Category(Row)
class RowCategory {
	/**
	 * 
	 * @param closure
	 * @return
	 */
	Object find(Closure<?> closure){
		int s = this.firstCellNum
		int e = this.lastCellNum
		Cell res = null;
		for(int i = s; i < e; i++){
			if(this.getCell(i) == null){
				continue
			}
			if(closure.call(this.getCell(i)) ){
				res = this.getCell(i)
				break
			}
		}
		return res
	}
	
	Object findAll(Closure<?> closure){
		int s = this.firstCellNum
		int e = this.lastCellNum
		List<Cell> res = [];
		for(int i = s; i < e; i++){
			if(this.getCell(i) == null){
				continue
			}
			if(closure.call(this.getCell(i)) ){
				res.add(this.getCell(i))
			}
		}
		return res
	}

	Cell getFirst(){
		return this?.getCell(this.firstCellNum)
	}

	Cell getLast(){
		return this?.getCell(this.lastCellNum - 1)
	}
}
