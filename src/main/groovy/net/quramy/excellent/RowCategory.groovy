package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author yosuke
 *
 */
@Category(Row)
class RowCategory {
	Cell getFirst(){
		return this?.getCell(this.firstCellNum)
	}

	Cell getLast(){
		return this?.getCell(this.lastCellNum - 1)
	}
}
