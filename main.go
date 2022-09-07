package main

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	t := time.Now()
	ymd := t.Format("20060102")
	filename1 := "../downloads/" + ymd + "oddlottrantest.xlsx"
	filename2 := "../downloads/2021092880111001test.xlsx"
	// ファイルを読み込み
	f1, _ := excelize.OpenFile(filename1)
	f2, _ := excelize.OpenFile(filename2)

	// A列に値を設定
	i := 1
	haihun := strconv.Quote("-")
	zero := strconv.Quote("0")
	for ; ; i++ {
		cellNo := "B" + strconv.Itoa(i)
		cellValue := f2.GetCellValue("Sheet2", cellNo)
		if cellValue != "" {
			s2CellA := "A" + strconv.Itoa(i)
			s2CellC := "C" + strconv.Itoa(i)
			s2CellD := "D" + strconv.Itoa(i)
			f2.SetCellFormula("Sheet2", s2CellA, s2CellC+"&"+haihun+"&"+s2CellD+"&"+zero)
		} else {
			i = i - 1
			break
		}
	}
	fmt.Println(i)

	// セルの設定
	for x := 2; x < i+2; x++ {
		cellNo1 := "C" + strconv.Itoa(x)
		cellNo1a := "A$" + strconv.Itoa(x)
		cellNo1b := "C$" + strconv.Itoa(x)
		cellNo2 := "F" + strconv.Itoa(x)
		cellNo3 := "N" + strconv.Itoa(x)
		cellNo4 := "O" + strconv.Itoa(x)
		cellNo5a := "H$" + strconv.Itoa(x)
		cellNo5b := "I$" + strconv.Itoa(x)
		cellNo6 := "L" + strconv.Itoa(x)
		cellNo8 := "E" + strconv.Itoa(x-1)
		c1 := f1.GetCellValue("Sheet1", cellNo1)
		c2 := f1.GetCellValue("Sheet1", cellNo2)
		c3 := f1.GetCellValue("Sheet1", cellNo3)
		c3a, _ := strconv.ParseFloat(c3, 64)
		c3g := ""
		if math.Floor(c3a) == c3a {
			c3b, _ := strconv.Atoi(c3)
			c3c := convert(c3b)
			c3d := c3c + ".0"
			c3g = c3d
			fmt.Println(c3g)
		} else {
			c3d := strings.Split(c3, ".")
			c3e, _ := strconv.Atoi(c3d[0])
			c3f := convert(c3e) + "." + c3d[1]
			c3g = c3f
			fmt.Println("整数ではない", c3g)
		}
		c4 := f1.GetCellValue("Sheet1", cellNo4)
		c4a, _ := strconv.Atoi(c4)
		c4b := convert(c4a)
		c6 := f1.GetCellValue("Sheet1", cellNo6)
		koza := f2.GetCellValue("Sheet2", cellNo8)
		c8 := 0
		if koza == "specific" {
			c8 = 1
		}
		cellNo10 := "D" + strconv.Itoa(x+13)
		cellNo11 := "E" + strconv.Itoa(x+13)
		cellNo12 := "F" + strconv.Itoa(x+13)
		cellNo13 := "G" + strconv.Itoa(x+13)
		cellNo14 := "H" + strconv.Itoa(x+13)
		cellNo15 := "I" + strconv.Itoa(x+13)
		cellNo16 := "J" + strconv.Itoa(x+13)
		cellNo17 := "K" + strconv.Itoa(x+13)
		cellNo18 := "L" + strconv.Itoa(x+13)
		cellNo19 := "M" + strconv.Itoa(x+13)
		cellNo20 := "N" + strconv.Itoa(x+13)
		f2.SetCellValue("hakabu", cellNo10, c1)
		f2.SetCellValue("hakabu", cellNo11, 1)
		f2.SetCellValue("hakabu", cellNo12, c2)
		f2.SetCellValue("hakabu", cellNo13, "")
		f2.SetCellValue("hakabu", cellNo14, c3g)
		f2.SetCellValue("hakabu", cellNo15, "09:00")
		f2.SetCellValue("hakabu", cellNo16, c4b)
		c5 := "[" + ymd + "oddlottrantest.xlsx]Sheet1!$"
		c5b := "Sheet1!A:B,2,FALSE)"
		f2.SetCellFormula("hakabu", cellNo17, c5+cellNo5a+"&"+haihun+"&"+c5+cellNo5b)
		c6a := strings.Split(c6, "-")
		c6b := "20" + c6a[2] + "/" + c6a[0] + "/" + c6a[1]
		f2.SetCellValue("hakabu", cellNo18, c6b)
		f2.SetCellFormula("hakabu", cellNo19, "VLOOKUP("+c5+cellNo1a+"&"+haihun+"&"+c5+cellNo1b+","+c5b)
		f2.SetCellValue("hakabu", cellNo20, c8)
	}

	if f2.WorkBook != nil && f2.WorkBook.CalcPr != nil {
		f2.WorkBook.CalcPr.FullCalcOnLoad = true
	}
	f2.SaveAs(filename2)

}

func convert(integer int) string {
	arr := strings.Split(fmt.Sprintf("%d", integer), "")
	cnt := len(arr) - 1
	res := ""
	i2 := 0
	for i := cnt; i >= 0; i-- {
		if i2 > 2 && i2%3 == 0 {
			res = fmt.Sprintf(",%s", res)
		}
		res = fmt.Sprintf("%s%s", arr[i], res)
		i2++
	}
	return res
}
