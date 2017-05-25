package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

// Row
const (
	iSheetID = 1

	iPaperID = 0
	//iPaperTitle   = 1
	iTrackID = 2
	//iTrackTile    = 3
	iSessionID = 4
	//iSessionTitle = 5
	iPaperType   = 25
	iUTCTime     = 29
	iConstraint1 = 30
	iConstraint2 = 31
	iConstraint3 = 32
)

// id papier | id track | pleinier | session | duree | utc time | 3 contraintes

func main() {
	excelFileName := "./data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Fatal(err)
	}

	outputFile, err := os.Create("result.txt")
	if err != nil {
		log.Fatal("Cannot create file", err)
	}
	defer outputFile.Close()

	cellsUsefull := []int{
		iPaperID, iTrackID, iSessionID,
		iPaperType, iUTCTime, iConstraint1, iConstraint2, iConstraint3,
	}
	sheet := xlFile.Sheets[iSheetID]
	for rowIndex, row := range sheet.Rows {
		if rowIndex != 0 {
			line := ""
			for cellIndex, cellID := range cellsUsefull {
				if cellID < len(row.Cells) {
					line += cellParser(cellID, row.Cells[cellID])
				}
				if cellIndex != len(cellsUsefull)-1 {
					line += "|"
				}
			}
			fmt.Fprintf(outputFile, "%s\n", line)
		}
	}

	// for _, sheet := range xlFile.Sheets {
	// 	for _, row := range sheet.Rows {
	// 		for _, cell := range row.Cells {
	// 			text, _ := cell.String()
	// 			fmt.Printf("%s\n", text)
	// 		}
	// 	}
	// }

}

func cellParser(cellID int, cell *xlsx.Cell) (res string) {
	switch cellID {
	case iPaperID:
		i, _ := cell.Int()
		res = strconv.Itoa(i)
	case iTrackID:
		i, _ := cell.Int()
		res = strconv.Itoa(i)
	case iSessionID:
		i, _ := cell.Int()
		res = strconv.Itoa(i)
	case iUTCTime:
		s, _ := cell.String()
		res = parseUTC(s)
	case iPaperType:
		s, _ := cell.String()
		res = strconv.Itoa(parseDuration(s))
	default:
		res, _ = cell.String()
	}

	return
}

func parseDuration(paperType string) (res int) {
	switch paperType {
	case "Tutorial":
		res = 120
	case "Plenary talk":
		res = 60
	case "Invited talk":
		res = 30
	case "Advanced Introduction invited talk":
		res = 30
	case "New Result invited paper":
		res = 30
	case "Full paper":
		res = 30
	case "Short paper":
		res = 15
	case "Poster":
		res = 5
	default:
		res = 0
		fmt.Println("Paper type not found: " + paperType)
	}

	return
}

func parseUTC(utc string) (res string) {
	res = strings.Replace(utc, "UTC", "", -1)
	res = strings.Replace(res, "UCT", "", -1)
	res = strings.Replace(res, "hours", "", -1)
	res = strings.Replace(res, "hour", "", -1)
	res = strings.Trim(res, " ")
	_, err := strconv.ParseFloat(res, 64)
	if err != nil {
		fmt.Println(err.Error())
		res = ""
	}
	return
}
