package net.quramy

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import net.quramy.excellent.Excellent;

Sheet sheet = Excellent.open("src/test/resources/test.xlsx").Messages

println "=" * 100
Cell headCell = sheet.findCell{it=~/head/}

println "The position of head cell is ${headCell.label}."// ->A4

// "up", "down", "left", "right" provide transition
println "The value of lower right head cell is ${headCell.down.right}."

Cell footCell = sheet.findCell{it=~/END/}
println "The position of foot cell is ${footCell.label}."

println "=" * 100

// Subtraction of 2 cells provides a CellRange.
(footCell.up - headCell.down.right).collect{row ->
	[id:row[1].value, detail:row[2].value]
}.findAll{it.id}.each { println "${it.id} : ${it.detail}" }