// usage: groovy modify.groovy read.xls result.xls

// ----------------------------
// Grape�ɂ�郉�C�u�����擾
@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
//groovy.grape.Grape.grab(group:'org.apache.poi', module:'poi', version:'3.0.2-FINAL')
//@Grab('org.apache.poi:poi:3.0.2-FINAL') // only for Groovy1.7

// ----------------------------
// Excel�t�@�C���̓ǂݍ���
def inputFile = new File(args[0])
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet1 = book.getSheetAt(0) // ��1�V�[�g
def sheet2 = book.getSheetAt(1) // ��2�V�[�g
def sheet3 = book.getSheet("Sheet3") // �V�[�g���Ŏw����\

def sheet = sheet1 // �ȍ~�Ŏg���V�[�g��I��
//def sheet = sheet2
//def sheet = sheet3

// ----------------------------
// �Z���̏�������
// HSSFCell�̓���͓ǂݍ��݂Ɠ����ł��邽�߁A�ȒP�̂��߃w���p���\�b�h�𗘗p����
def cell = { row, col ->
    sheet.getRow(row)?.getCell((short) col)
}

// ���O�̒l�m�F
println cell(0, 0).stringCellValue // A1
println cell(1, 0).stringCellValue // A2
println cell(2, 0).numericCellValue.intValue() // A3 (double->int)
println cell(3, 0).dateCellValue // A4
println cell(4, 0).booleanCellValue // A5

// ��������
cell(0, 0).setCellValue("Modified_A1") // A1
cell(1, 0).setCellValue("�ύX����_A2") // A2
cell(2, 0).setCellValue(7890) // A3
cell(3, 0).setCellValue(new Date()) // A4
cell(4, 0).setCellValue(false) // A5

// ����̒l�m�F
println cell(0, 0).stringCellValue // A1
println cell(1, 0).stringCellValue // A2
println cell(2, 0).numericCellValue.intValue() // A3 (double->int)
println cell(3, 0).dateCellValue // A4
println cell(4, 0).booleanCellValue // A5

// ----------------------------
// �V�KExcel�t�@�C���ւ̏o��
def outputFile = new File(args[1])
outputFile.withOutputStream { out ->
    book.write(out)
}

