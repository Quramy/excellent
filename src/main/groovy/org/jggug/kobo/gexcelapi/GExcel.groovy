/*
 * Copyright 2009 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.jggug.kobo.gexcelapi

import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
import org.apache.poi.ss.usermodel.*
import org.jggug.kobo.gexcelapi.CellLabelUtils as CLU
import java.lang.IndexOutOfBoundsException as IOOBEx

class GExcel {

    static {
        expandHSSFWorkbook()
        expandHSSFSheet()
        expandHSSFRow()
        expandHSSFCell()
    }

    private static expandHSSFWorkbook() {
        HSSFWorkbook.metaClass.define {
            getAt { int idx -> delegate.getSheetAt(idx) }
            getProperty { String name -> delegate.getSheet(name) }
        }
    }

    private static expandHSSFSheet() {
        def methods = {
            getProperty { name ->
                if (name == "rows") { return rows() }
                if (name ==~ /_\d+/) { return delegate.getRow(CLU.rowIndex(name)) }
                if (name ==~ /[a-zA-Z]+_/) {
                    int columnIndex = CLU.columnIndex(name)
                    return new CellRange(delegate, delegate.getFirstRowNum(), columnIndex, delegate.getLastRowNum(), columnIndex)
                }
                if (name ==~ /[a-zA-Z]+\d+/) {
                    try { return delegate.getRow(CLU.rowIndex(name))?.getCell(CLU.columnIndex(name)) } catch (IOOBEx e) { return null }
                }
                if (name ==~ /([a-zA-Z]+\d+)_([a-zA-Z]+\d+)/) {
                    def token = name.split("_")
                    return new CellRange(delegate, token[0], token[1])
                }
                null
            }
            setProperty { name, value ->
                if (name ==~ /[a-zA-Z]+\d+/) {
                    try { delegate.getRow(CLU.rowIndex(name))?.getCell(CLU.columnIndex(name)).setCellValue(value) } catch (IOOBEx e) { return null }
                }
            }
            rows { delegate?.findAll{true} }
            validate { delegate.rows.every { row -> row.validate() } }
        }
        HSSFSheet.metaClass.define methods
    }

    private static expandHSSFRow() {
        def methods = {
            getAt { int idx -> delegate.getCell(idx) }
            getProperty { name ->
                if (name ==~ /[a-zA-Z]+_/) { return delegate.getCell(CLU.columnIndex(name)) }
                null
            }
            validate { delegate.every { cell -> cell.validate() } }
        }
        HSSFRow.metaClass.define methods
    }

    private static expandHSSFCell() {
        HSSFCell.metaClass.__validators__ = null
        HSSFCell.metaClass.define {
            isStringType  { delegate.cellType == Cell.CELL_TYPE_STRING }
            isNumericType { delegate.cellType == Cell.CELL_TYPE_NUMERIC }
            isBooleanType { delegate.cellType == Cell.CELL_TYPE_BOOLEAN }
            getValue {
                // implicitly accessing value by appropriate type
                switch(delegate.cellType) {
                    case Cell.CELL_TYPE_STRING:  return delegate.stringCellValue
                    case Cell.CELL_TYPE_NUMERIC: return delegate.numericCellValue
                    case Cell.CELL_TYPE_BOOLEAN: return delegate.booleanCellValue
                    default: throw new RuntimeException("unsupported cell type: ${delegate.cellType}")
                }
            }
            setValue { value -> delegate.setCellValue(value) }
            leftShift { value -> delegate.setCellValue(value) }
            asType { Class type ->
                // explicitly accessing value by appropriate type
                switch(type) {
                    case Double:  return delegate.numericCellValue
                    case Integer: return delegate.numericCellValue.intValue()
                    case Boolean: return delegate.booleanCellValue
                    case Date:    return delegate.dateCellValue
                    case String:  return new String(delegate.stringCellValue.getBytes("Windows-31J")) // Windows-31J -> default encoding
                    default: throw new RuntimeException("unsupported cell type: ${delegate.cellType}")
                }
            }
            getLabel { CLU.cellLabel(delegate.rowIndex, delegate.columnIndex) }
            clearValidators {
                delegate.__validators__ = []
            }
            getValidators {
                if (delegate.__validators__ == null) { delegate.clearValidators() }
                delegate.__validators__
            }
            addValidator { Closure validator ->
                delegate.validators << validator
            }
            setValidator { Closure validator ->
                delegate.clearValidators()
                delegate.validators << validator
            }
            validate {
                delegate.validators.every { validator -> validator.call(delegate) }
            }
        }
    }

    static open(String file) { open(new File(file)) }
    static open(File file) { open(new FileInputStream(file)) }
    static open(InputStream is) { new HSSFWorkbook(new POIFSFileSystem(is)) }
}
