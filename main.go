package main

import (
	"errors"
	"fmt"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

// Row
const (
	/*
		iFirstSheetID = 1
		iLastSheetID  = 1

		iPaperID = 0
		iTrackID = 2
		iSessionID = 4
		iPaperType   = 25
		iUTCTime     = 29
		iConstraint1 = 30
		iConstraint2 = 31
		iConstraint3 = 32
	*/

	iFirstSheetID = 2
	iLastSheetID  = 13

	iPaperID     = 0
	iTrackID     = 2
	iSessionID   = 4
	iPaperType   = 25
	iUTCTime     = 32
	iConstraint1 = 33
	iConstraint2 = 34
	iConstraint3 = 35
)

// id papier | id track | pleinier | session | duree | utc time | 3 contraintes

func main() {
	excelFileName := "./data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Fatal(err)
	}

	cellsUsefull := []int{
		iPaperID, iTrackID, iSessionID,
		iPaperType, iUTCTime, iConstraint1, iConstraint2, iConstraint3,
	}
	lines := make([]string, 1)
	consCounter := make([]int, 3)

	for sheetIndex := iFirstSheetID; sheetIndex <= iLastSheetID; sheetIndex++ {
		sheet := xlFile.Sheets[sheetIndex]
		for rowIndex, row := range sheet.Rows {
			if len(row.Cells) > 0 {
				paperID, _ := row.Cells[iPaperID].Int()
				if rowIndex != 0 && paperID >= 0 {
					line := ""
					for cellIndex, cellID := range cellsUsefull {
						if cellID < len(row.Cells) {
							data, err := cellParser(cellID, row.Cells[cellID], &consCounter)
							if err != nil {
								fmt.Printf("Error in sheet %s at line %d :\n%s\n", sheet.Name, rowIndex+1, err)
							}
							line += data
						}
						if cellIndex != len(cellsUsefull)-1 {
							line += "|"
						}
					}
					lines = append(lines, line)
				}
			}
		}
	}

	//lines[0] = fmt.Sprintf("%d|%d|%d", consCounter[0], consCounter[1], consCounter[2])

	// Write file
	outputFile, err := os.Create("result.txt")
	if err != nil {
		log.Fatal("Cannot create file", err)
	}
	defer outputFile.Close()
	for _, line := range lines {
		fmt.Fprintf(outputFile, "%s\n", line)
	}
}

func cellParser(cellID int, cell *xlsx.Cell, consCounter *[]int) (string, error) {
	var res string
	var err error
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
		res, _ = cell.String()
		//res = parseUTC(s)
	case iPaperType:
		s, _ := cell.String()
		res = strconv.Itoa(parseDuration(s))
	case iConstraint1:
		s, _ := cell.String()
		res, err = parseConstraint(s)
		if res != "" {
			(*consCounter)[0]++
		}
	case iConstraint2:
		s, _ := cell.String()
		res, err = parseConstraint(s)
		if res != "" {
			(*consCounter)[1]++
		}
	case iConstraint3:
		s, _ := cell.String()
		res, err = parseConstraint(s)
		if res != "" {
			(*consCounter)[2]++
		}

	default:
		res, _ = cell.String()
	}

	return res, err
}

func checkConstraint(cons string) error {
	resultReg, _ := regexp.Compile(`\[(\d*):(\d*),(\d*):(\d*)\]`)
	matches := resultReg.FindAllStringSubmatch(cons, -1)

	for _, match := range matches {
		v1, err1 := strconv.Atoi(match[1])
		v2, err2 := strconv.Atoi(match[3])
		if err1 != nil || err2 != nil || v1 > v2 {
			return errors.New("Can't parse the constraint: " + cons + "\nThe correct format is [hh:mm, hh:mm] or [hh, hh]\n")
		}
	}
	return nil
}

func replaceConstraint(cons string) string {
	resultReg, _ := regexp.Compile(`24:00`)
	matches := resultReg.FindAllStringSubmatch(cons, -1)
	res := resultReg.ReplaceAllLiteralString(cons, "23:59")
	if len(matches) > 0 {
		//fmt.Println(cons + " replaced by" + res)
	}
	return res
}

func parseConstraint(cons string) (string, error) {
	res := ""
	cons = strings.Trim(cons, " ")
	var err error
	valid := false

	resultReg, _ := regexp.Compile(`\[(\d*):(\d*),(\d*):(\d*)\]`)
	matches := resultReg.FindAllStringSubmatch(cons, -1)
	if len(matches) > 0 || cons == "" {
		res = cons
		valid = true
	}

	if !valid {
		re1, _ := regexp.Compile(`\[(\d*),(\d*)\]`)
		matches = re1.FindAllStringSubmatch(cons, -1)
		for i, match := range matches {
			res += fmt.Sprintf("[%s:00,%s:00]", match[1], match[2])
			if i != len(matches)-1 {
				res += ","
			}
		}
		matches = resultReg.FindAllStringSubmatch(res, -1)
		if len(matches) <= 0 {
			err = errors.New("Can't parse the constraint: " + cons + "\nThe correct format is [hh:mm, hh:mm] or [hh, hh]\n")
		}
	}

	res = replaceConstraint(res)
	if err == nil {
		err = checkConstraint(res)
	}
	return res, err
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
	case "Young researcher":
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
