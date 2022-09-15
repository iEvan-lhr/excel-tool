package number

// CountNumber 目前最大支持下标为ZZ
func CountNumber(str string) int {
	if str == "" {
		return 0
	} else if len(str) == 1 {
		return int(str[0] - 'A')
	} else if len(str) == 2 {
		return (int(str[0]-'A')+1)*26 + int(str[1]-'A')
	} //sum := 0
	//values := make([]int, len(str))
	//for i, v := range str {
	//	if v >= 'A' && v <= 'Z' {
	//		values[i] = int(v-'A') + 1
	//	} else {
	//		panic("Please Check Str")
	//	}
	//}
	//for _, v := range values {
	//
	//}
	return -1
}
