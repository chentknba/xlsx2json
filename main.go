package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"encoding/json"
	"github.com/tealeg/xlsx"
)

var cfg = map[string]string{}

func loadConf() error {
	bytes, err := ioutil.ReadFile("conf.json")
	if err != nil {
		return err
	}

	if err := json.Unmarshal(bytes, &cfg); err != nil {
		return err
	}

	return nil
}

func main() {
	if err := loadConf(); err != nil {
		fmt.Println("load config err: %v", err)
		return
	}

	execl_path := cfg["execl_path"]
	dict_path := cfg["dict_path"]

	fmt.Printf("配置表路径: %v\n", execl_path)
	fmt.Printf("生成json路径: %v\n", dict_path)
	fmt.Printf("execl: %v\n", cfg["file"])

	file := execl_path + "test.xlsx"
	out  := dict_path + "test.json"

	xlfile, err := xlsx.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}

	// only deal with Sheet 0
	sh := xlfile.Sheets[0]
	if sh == nil {
		fmt.Println("cant find sheet, must something wrong happen.")
		return
	}

	nrows := len(sh.Rows)
	if nrows < 3 {
		fmt.Println("wrong fmt.")
		return
	}

	// row 0: identity
	// row 1: type declare
	// row 2: comment
	// row 3: data

	lbrace, rbrace := "{", "}"
	lbraket, rbraket := "[", "]"

	var json_content string
	json_start := true
	json_content = lbraket

	rows := sh.Rows

	for ri, row := range rows {
		if ri >= 3 {
			var object string

			object = lbrace
			object += "\r\n"

			begin := true

			for ci, cell := range row.Cells {
				var item string
				k, _ := sh.Cell(0, ci).String()
				t, _ := sh.Cell(1, ci).String()

				text, _ := cell.String()

				k = "\"" + k + "\""

				if t == "string" {
					text = "\"" + text + "\""
				}

				item = k + ":" + text

				if !begin {
					object += ","
					object += "\r\n"

				}else {
					begin = false
				}

				object += "\t"
				object += item
			}

			object += "\r\n"
			object += rbrace

			if json_start {
				json_content += "\r\n"

				json_start = false
			}else {
				json_content += ","
				json_content += "\r\n"
			}

			json_content += "\t"
			json_content += object
		}
	}

	json_content += "\r\n"
	json_content += rbraket

	f, err := os.OpenFile(out, os.O_RDWR | os.O_CREATE, 0755)
	if err != nil {
		fmt.Println("open out file fail.")
		return
	}

	defer f.Close()

	f.WriteString(json_content)

	//fmt.Println("json content :%v", json_content)
}
