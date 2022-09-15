package file

import "github.com/xuri/excelize/v2"

func DeleteNoneData(name, sheet string, del [][]int, mDel map[string]int) {
	file, err := excelize.OpenFile(name)
	if err != nil {
		panic(err)
	}
	for _, v := range del {
		line := v[1]
		removeRows(file, sheet, v[0], line)
	}
	rows, err := file.GetRows(sheet)
	if err != nil {
		panic(err)
	}
	for i := 0; i < len(rows); i++ {
		for _, s := range rows[i] {
			if v, ok := mDel[s]; ok {
				removeRows(file, sheet, i+1, v)
				rows = append(rows[:i], rows[i+v:]...)
				i -= v
				break
			}
		}
	}
	err = file.SetSheetViewOptions(sheet, -1, excelize.ShowGridLines(false))
	if err != nil {
		panic(err)
	}
	err = file.Save()
	if err != nil {
		panic(err)
	}
}

func removeRows(file *excelize.File, sheet string, start, line int) {
	for i := 0; i < line; i++ {
		err := file.RemoveRow(sheet, start)
		if err != nil {
			panic(err)
		}
	}
}
