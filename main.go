package main

import (
	"encoding/json"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"sync"
)

var cfg = map[string]string{}
var dict_cfg = map[string]string{}

var execl_path string
var desc_json string
var dict_path string

var wg sync.WaitGroup

func loadConf() error {
	bytes, err := ioutil.ReadFile("conf.json")
	if err != nil {
		return err
	}

	if err := json.Unmarshal(bytes, &cfg); err != nil {
		return err
	}

	execl_path = cfg["execl_path"]
	desc_json = execl_path + "desc.json"

	bytes, err = ioutil.ReadFile(desc_json)
	if err != nil {
		return err
	}

	if err := json.Unmarshal(bytes, &dict_cfg); err != nil {
		return err
	}

	dict_path = cfg["dict_path"]

	return nil
}

func gen(exel_name, dict_name string) {
	defer wg.Done()

	file := execl_path + exel_name + ".xlsx"
	out := dict_path + dict_name + ".json"

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

	// fmt.Printf("dict: %v, rows: %v\n", dict_name, nrows)

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
			end := true

			for _, cell := range row.Cells {
				if str, _ := cell.String(); str != "" {
					end = false
					break
				}
			}

			if end {
				// fmt.Printf("%v end.\n", dict_name)
				break
			}

			for ci, cell := range row.Cells {
				var item string
				k, _ := sh.Cell(0, ci).String()
				if k == "" {
					continue
				}

				t, _ := sh.Cell(1, ci).String()

				text, _ := cell.String()

				if text == "" {
					if k == "int" {
						text = "0"
					} else if k == "float" {
						text = "0.0"
					}
				}

				k = "\"" + k + "\""

				if t == "string" {
					text = "\"" + text + "\""
				}

				item = k + ":" + text

				if !begin {
					object += ","
					object += "\r\n"

				} else {
					begin = false
				}

				object += "\t\t"
				object += item
			}

			object += "\r\n"
			object += "\t"
			object += rbrace

			if json_start {
				json_content += "\r\n"

				json_start = false
			} else {
				json_content += ","
				json_content += "\r\n"
			}

			json_content += "\t"
			json_content += object
		}
	}

	json_content += "\r\n"
	json_content += rbraket

	f, err := os.OpenFile(out, os.O_RDWR|os.O_CREATE, 0755)
	if err != nil {
		fmt.Println("open out file fail.")
		return
	}

	defer f.Close()

	f.WriteString(json_content)

	fmt.Printf("generate %v, rows %v\n", dict_name, nrows)

}

func main() {
	if err := loadConf(); err != nil {
		fmt.Println("load config err: %v", err)
		return
	}

	fmt.Printf("配置表路径: %v\n", execl_path)
	fmt.Printf("生成json路径: %v\n", dict_path)
	fmt.Printf("execl: %v\n", cfg["file"])

	for dict_name, exel_name := range dict_cfg {
		wg.Add(1)

		go gen(exel_name, dict_name)
	}

	wg.Wait()
}
