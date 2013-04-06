package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Expand Cell.
 * Provide transition methods(up, down, left, right).
 * 
 * @author Quramy
 *
 */
@Category(Cell)
class CellCategory {

	Cell getLeft(){
		return getHorizontal(this, -1)
	}

	Cell getRight(){
		return getHorizontal(this, 1)
	}

	Cell getUp(){
		return getVertical(this, -1)
	}

	Cell getDown(){
		return getVertical(this, 1)
	}

	private static Cell getHorizontal(Cell ref, int offset){
		if(ref == null){
			return null
		}
		if(ref.columnIndex + offset >= 0){
			return ref.row.getCell(ref.columnIndex + offset, Row.CREATE_NULL_AS_BLANK)
		}else{
			return ref
		}
	}

	private static Cell getVertical(Cell ref, int offset){
		if(ref == null){
			return null
		}
		int pos = ref.rowIndex + offset
		if(pos < 0){
			return ref
		}
		Row refRow = ref.sheet.getRow(pos) ?: ref.sheet.createRow(pos)
		return refRow.getCell(ref.columnIndex, Row.CREATE_NULL_AS_BLANK)
	}
}
