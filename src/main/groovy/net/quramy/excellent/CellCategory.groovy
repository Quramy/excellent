package net.quramy.excellent

import org.apache.poi.ss.usermodel.Cell;

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
		if(this.columnIndex > 0){
			return this.row.getCell(this.columnIndex - 1)
		}else{
			return this
		}
	}

	Cell getRight(){
		return this.row.getCell(this.columnIndex + 1)
	}

	Cell getUp(){
		if(this.rowIndex  > 0){
			return this.row.sheet.getRow(this.rowIndex - 1).getCell(this.columnIndex)
		}else{
			return this
		}
	}

	Cell getDown(){
		return this.row.sheet.getRow(this.rowIndex + 1).getCell(this.columnIndex)
	}
}
