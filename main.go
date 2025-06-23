package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

var Wotsv2YamlFiles = []string{
	"D:\\Project-Wotsv2\\wotsv2-server\\sheet\\sheet-gamedata.yaml",
	"D:\\Project-Wotsv2\\wotsv2-server\\sheet\\sheet-localization.yaml",
}
var Wotsv2ExcelDir = []string{
	"D:\\Project-Wotsv2\\wotsv2-client\\Assets\\ExcelExportMaker\\Excels\\GameData",
	"D:\\Project-Wotsv2\\wotsv2-client\\Assets\\ExcelExportMaker\\Excels\\Localization",
}
var CoffeeExcelDir = []string{
	"D:\\Project-CoffeeAgent\\coffeeagent-client\\Assets\\ExcelExportMaker\\Excels\\GameData",
	"D:\\Project-CoffeeAgent\\coffeeagent-client\\Assets\\ExcelExportMaker\\Excels\\Localization",
}

//TIP <p>To run your code, right-click the code and select <b>Run</b>.</p> <p>Alternatively, click
// the <icon src="AllIcons.Actions.Execute"/> icon in the gutter and select the <b>Run</b> menu item from here.</p>

func main() {
	// yaml1 := ReadYaml(YAML_File1)
	// yaml2 = ReadYaml(YAML_File2)

	excel := []*Excel{}

	for _, dir := range Wotsv2ExcelDir {
		files := FindExcel(dir)
		for _, file := range files {
			excel = append(excel, ReadFile(file))
		}
	}

	fmt.Printf("excel[0] ==> %s\n", func() []byte {
		jsonString, _ := json.MarshalIndent(excel[0], "", "    ")
		return jsonString
	}())

	ExportSchema("wotsv2-sheet-schema.txt", excel...)

	excel = []*Excel{}

	for _, dir := range CoffeeExcelDir {
		files := FindExcel(dir)
		for _, file := range files {
			excel = append(excel, ReadFile(file))
		}
	}

	ExportSchema("coffee-sheet-schema.txt", excel...)

}

func ReadYaml(file string) map[string][]string {
	f, err := os.Open(file)

	if err != nil {
		panic(err)
	}

	defer f.Close()

	scanner := bufio.NewScanner(f)

	yaml := map[string][]string{}
	key := ""

	for scanner.Scan() {
		text := scanner.Text()
		re1 := regexp.MustCompile("([a-zA-z]+):$")
		re2 := regexp.MustCompile("^.*:.*$")

		if re1.MatchString(text) {
			match := re1.FindStringSubmatch(text)
			key = match[1]
		} else if re2.MatchString(text) {
			parts := strings.Split(text, ":")
			yaml[parts[0]] = []string{strings.TrimSpace(parts[1])}
		} else {
			text = strings.TrimPrefix(text, "  - ")
			yaml[key] = append(yaml[key], text)
		}
	}

	if err = scanner.Err(); err != nil {
		panic(err)
	}

	return yaml
}

func FindExcel(dir string) []string {
	files, err := os.ReadDir(dir)

	if err != nil {
		return nil
	}

	var result []string

	for _, file := range files {
		if !file.IsDir() &&
			strings.HasSuffix(file.Name(), ".xlsx") &&
			strings.HasPrefix(file.Name(), "~") == false {
			result = append(result, dir+"\\"+file.Name())
		}
	}

	return result
}

func ReadFile(file string) *Excel {
	if _, err := os.Stat(file); os.IsNotExist(err) {
		log.Fatalf("找不到檔案: %s", file)
		return nil
	}

	excel := &Excel{
		Dir:  filepath.Base(filepath.Dir(file)),
		File: file,
		Name: filepath.Base(file),
	}

	f, err := excelize.OpenFile(file)

	if err != nil {
		log.Fatal(err)
	}

	defer f.Close()

	for index := 0; index < f.SheetCount; index++ {
		// 取得第一個工作表名稱
		sheetName := f.GetSheetName(index)

		// 讀取指定工作表所有資料
		rows, err := f.GetRows(sheetName)

		if err != nil {
			log.Fatal(err)
		}

		fieldCount := len(rows[0])
		table := &Table{
			Name: sheetName,
		}

		for i := 1; i < fieldCount; i++ {
			field := &Field{
				Visible: getCell(rows, 0, i, ""),
				Type:    getCell(rows, 1, i, ""),
				Key:     getCell(rows, 2, i, ""),
				Name:    getCell(rows, 3, i, ""),
			}
			table.Fields = append(table.Fields, field)
		}

		excel.Tables = append(excel.Tables, table)
	}

	return excel
}

func ExportSchema(exportFile string, excel ...*Excel) bool {
	content := ""

	for _, itor := range excel {
		content += fmt.Sprintf("[%v/%v]\n", itor.Dir, itor.Name)

		for _, table := range itor.Tables {
			content += fmt.Sprintf("    - %v:\n", table.Name)

			for _, field := range table.Fields {
				if field.Visible == "Ignore" || field.Key == "" {
					//content += fmt.Sprintf("          (ignore) %v (%v) - %v\n", field.Key, field.Type, field.Name)
					continue
				}

				content += fmt.Sprintf("        * %v (%v) - %v\n", field.Key, field.Type, field.Name)
			}
		}
	}

	err := os.WriteFile(exportFile, []byte(content), 0644)

	if err != nil {
		panic(err)
	}

	return true
}

func getCell(rows [][]string, row int, col int, defaultVal string) string {
	if row < len(rows) && col < len(rows[row]) {
		return rows[row][col]
	}

	return defaultVal
}

type Excel struct {
	Dir    string
	File   string
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
