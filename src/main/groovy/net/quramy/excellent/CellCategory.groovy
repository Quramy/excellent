package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.jggug.kobo.gexcelapi.CellRange

/**
 * Expand Cell.
 * Provide transition methods(up, down, left, right, + x, -x, Cell a - Cell b).
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

	Cell plus(int offset){
		return getHorizontal(this, offset)
	}

	Cell minus(int offset){
		return getHorizontal(this, -offset)
	}

	CellRange minus(Cell offset){
		Cell a = this, b = offset
		if(!a || !b || a.sheet != b.sheet){
			return null
		}

		int top, bottom, left, right
		top = Math.min(a.rowIndex, b.rowIndex)
		bottom = Math.max(a.rowIndex, b.rowIndex)
		left = Math.min(a.columnIndex, b.columnIndex)
		right = Math.max(a.columnIndex, b.columnIndex)

		return new CellRange(this.sheet, top, left, bottom, right)
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
