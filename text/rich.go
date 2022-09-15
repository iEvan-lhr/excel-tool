package text

import (
	"errors"
	"excel-tool/number"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

type Text struct {
	HighlightFont *excelize.Font `json:"highlight_font"`
	DefaultFont   *excelize.Font `json:"default_font"`
}

// SetHighlightText 使用占位符的方式高亮文本 使用空的str来使所有文字高亮 使用YES使文字使用默认格式高亮
func SetHighlightText(file *excelize.File, sheet, str string, placeholder, cell string, values []string, style Text) error {
	if style.DefaultFont == nil {
		style.DefaultFont = &excelize.Font{
			Color:  "#000000",
			Family: "Calibri",
			Size:   11,
		}
		err := file.SetCellRichText(sheet, cell, highlightsColor(str, placeholder, values, style))
		if err != nil {
			panic(err)
		}
	} else if style.HighlightFont == nil {
		return errors.New("Not HighlightFont Please Check ")
	}
	return nil
}

func DisplaceValues(name, sheet, row string, highlight int, values map[string]interface{}, style Text, placeholder string) {
	file, err := excelize.OpenFile(name)
	if err != nil {
		panic(err)
	}
	index := number.CountNumber(row)
	rows, err := file.GetRows(sheet)
	if err != nil {
		panic(err)
	}
	for i := 0; i < len(rows); i++ {
		if v, ok := values[rows[i][index]]; ok {
			switch v.(type) {
			case string:
				if highlight == 0 {
					err = file.SetCellValue(sheet, row+strconv.Itoa(i+1), v)
					if err != nil {
						panic(err)
					}
				} else {
					err = SetHighlightText(file, sheet, v.(string), placeholder, row+strconv.Itoa(i+1), []string{v.(string)}, style)
					if err != nil {
						panic(err)
					}
				}
			case []string:
				for k, s := range v.([]string) {
					if k != 0 {
						rows = append(rows[:i], append([][]string{{"0"}}, rows[i:]...)...) // 在第i个位置插入1
						i++
						err = file.InsertRow(sheet, i+1)
						if err != nil {
							panic(err)
						}
					}
					if highlight == 0 {
						err = file.SetCellValue(sheet, row+strconv.Itoa(i+1), s)
						if err != nil {
							panic(err)
						}
					} else {
						err = SetHighlightText(file, sheet, s, placeholder, row+strconv.Itoa(i+1), []string{s}, style)
						if err != nil {
							panic(err)
						}
					}
				}
			}

		}

	}
	err = file.Save()
	if err != nil {
		panic(err)
	}
}

func highlightsColor(str, placeholder string, values []string, style Text) []excelize.RichTextRun {
	text := strings.Split(str, placeholder)
	var ans []excelize.RichTextRun
	i := 0
	if str == "" {
		return []excelize.RichTextRun{{
			Text: values[i],
			Font: style.HighlightFont}}
	} else if values[0] == "Yes" {
		return []excelize.RichTextRun{{
			Text: str,
			Font: style.DefaultFont}}
	} else if len(text) == 1 && text[0][0] == ' ' {
		ans = append(ans, excelize.RichTextRun{
			Text: values[i],
			Font: style.HighlightFont})
		ans = append(ans, excelize.RichTextRun{
			Text: text[0],
			Font: style.DefaultFont})
	} else {
		for _, s := range text {
			ans = append(ans, excelize.RichTextRun{
				Text: s,
				Font: style.DefaultFont})
			if i < len(values) {
				ans = append(ans, excelize.RichTextRun{
					Text: values[i],
					Font: style.HighlightFont})
			}
			i++
		}
	}
	return ans
}
