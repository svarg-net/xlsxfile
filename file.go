package xlsxfile

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
	"time"
)

func Open(nameFile string) excelize.File {
	now := time.Now()
	nameFile = nameFile + "_" + now.Format("2006-02-01") + ".xlsx"
	f, err := excelize.OpenFile(nameFile)
	if err != nil {
		f = excelize.NewFile()
		if err := f.SaveAs(nameFile); err != nil {
			fmt.Println("addRowToFile1:", err)
		}
		f, err = excelize.OpenFile(nameFile)
	}
	return *f
}

func Set(f *excelize.File, set []interface{}) (excelize.File, error) {

	rmax := lastRow(f)
	f.SetSheetRow("Sheet1", "A"+strconv.Itoa(rmax+1), &set)
	return *f, nil

}

func SetAll(f *excelize.File, set [][]interface{}) (excelize.File, error) {

	rmax := lastRow(f)
	for _, s := range set {
		rmax++
		f.SetSheetRow("Sheet1", "A"+strconv.Itoa(rmax), &s)
	}
	return *f, nil

}
func Save(file excelize.File) error {
	err := file.Save()
	if err != nil {
		fmt.Println("addRowToFile2:", err)
		return err
	}
	return nil
}

func lastRow(xlFile *excelize.File) int {
	rmax := len(xlFile.GetRows("Sheet1"))
	return rmax
}
