package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

//TIP <p>To run your code, right-click the code and select <b>Run</b>.</p> <p>Alternatively, click
// the <icon src="AllIcons.Actions.Execute"/> icon in the gutter and select the <b>Run</b> menu item from here.</p>

func main() {
	file := "D:\\Project-Wotsv2\\wotsv2-client\\Assets\\ExcelExportMaker\\Excels\\GameData\\stage.xlsx"

	excel := ReadFile(file)

	fmt.Printf("excel ==> %s\n", func() []byte {
		jsonString, _ := json.MarshalIndent(excel, "", "    ")
		return jsonString
	}())
}

func ReadFile(file string) *Excel {
	if _, err := os.Stat(file); os.IsNotExist(err) {
		log.Fatalf("找不到檔案: %s", file)
		return nil
	}

	excel := &Excel{
		Name: file,
	}

	f, err := excelize.OpenFile(file)

	if err != nil {
		log.Fatal(err)
	}

	defer f.Close()

	fmt.Println("f.SheetCount", f.SheetCount)

	for index := 0; index < f.SheetCount; index++ {
		// 取得第一個工作表名稱
		sheetName := f.GetSheetName(index)

		fmt.Println("sheetName => ", sheetName)

		// 讀取指定工作表所有資料
		rows, err := f.GetRows(sheetName)

		if err != nil {
			log.Fatal(err)
		}

		fieldCount := len(rows[0])

		fmt.Printf("rows[0] ==> %s\n", func() []byte {
			jsonString, _ := json.MarshalIndent(rows[0], "", "    ")
			return jsonString
		}())

		fmt.Println("fieldCount => ", fieldCount)

		table := &Table{
			Name: sheetName,
		}

		for i := 0; i < fieldCount; i++ {
			field := &Field{
				Visible: getCell(rows, 0, i, ""), // 預設 "false"
				Type:    getCell(rows, 1, i, ""), // 預設 "string"
				Key:     getCell(rows, 2, i, ""),
				Name:    getCell(rows, 3, i, ""),
			}
			fmt.Printf("field ==> %s\n", func() []byte {
				jsonString, _ := json.MarshalIndent(field, "", "    ")
				return jsonString
			}())

			table.Fields = append(table.Fields, field)
		}

		excel.Tables = append(excel.Tables, table)
	}

	return excel
}

func getCell(rows [][]string, row int, col int, defaultVal string) string {
	if row < len(rows) && col < len(rows[row]) {
		return rows[row][col]
	}
	return defaultVal
}

type Excel struct {
	Name   string
	Tables []*Table
}

type Table struct {
	Name   string
	Fields []*Field
}

type Field struct {
	Visible string
	Type    string
	Key     string
	Name    string
}
