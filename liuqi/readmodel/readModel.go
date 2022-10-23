package readmodel

import (
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type CpcRow struct {
	//出租方
	lessor string
	//日期
	date string
	//期限
	term string
	//服务费
	charge string
}

func ReadExcel(path string) {

	f, err := excelize.OpenFile(path)
	if err != nil {
		panic(err)
	}

	cpcRowDataList := ReadCpcSheet(f)
	wycfwfHeader := ReadWycfwfHeader(f)
	FormateWycfwfData(f, cpcRowDataList, wycfwfHeader)

	defer func() {
		if err := f.Close(); err != nil {
			panic(err)
		}
	}()
}

// 读取城配车数据
func ReadCpcSheet(file *excelize.File) map[int]CpcRow {
	sheetName := "城配车"
	rows, err := file.GetRows(sheetName)
	if err != nil {
		panic(err)
	}
	var (
		cpcRowListMap = make(map[int]CpcRow)
		lessorNum     int
		dateNum       int
		termNum       int
		chargeNum     int
		cpcRowNum     int = 0
	)
	for rIndex, row := range rows {
		if rIndex == 0 {
			for cIndex, colStr := range row {
				if colStr == "出租方" {
					lessorNum = cIndex
				} else if colStr == "签约日期" {
					dateNum = cIndex
				} else if colStr == "租赁期限" {
					termNum = cIndex
				} else if colStr == "橙电服务费" {
					chargeNum = cIndex
				}
			}
		} else {
			var tmpRow CpcRow
			for cIndex, colStr := range row {
				if cIndex == lessorNum {
					tmpRow.lessor = colStr
				} else if cIndex == dateNum {
					dateArr := strings.Split(colStr, "/")
					if len(dateArr) == 3 {
						dateStr := dateArr[0] + "年" + dateArr[1] + "月"
						tmpRow.date = dateStr
					} else {
						continue
					}
				} else if cIndex == termNum {
					tmpRow.term = colStr
				} else if cIndex == chargeNum {
					tmpRow.charge = colStr
				}
			}
			cpcRowListMap[cpcRowNum] = tmpRow
			cpcRowNum++
		}
	}
	return cpcRowListMap
}

// 网约车服务费表头
func ReadWycfwfHeader(file *excelize.File) map[string]interface{} {
	rows, err := file.GetRows("网约车运营服务费")
	if err != nil {
		panic(err)
	}
	var companyStartColNumMap = make(map[string]int)
	for rIndex, row := range rows {
		if rIndex == 1 {
			for cIndex, colStr := range row {
				if colStr == "当期合计" {
					break
				}
				if colStr != "" {
					companyStartColNumMap[colStr] = cIndex

				}
			}
		}
	}
	var res = make(map[string]interface{})
	res["startRowNum"] = len(rows)
	res["companyStartColNumMap"] = companyStartColNumMap
	return res
}

type WycRow struct {
	qishu   string
	taishu  int
	fee     string
	company string
	qijian  string
}

// 整理网约车服务费数据
func FormateWycfwfData(file *excelize.File, cpcRowDataList map[int]CpcRow, wycfwfHeader map[string]interface{}) {
	companyStartColNumMap := wycfwfHeader["companyStartColNumMap"]
	var (
		companyGroup               = make(map[string]string)
		companyShortToLongName     = make(map[string]string)
		cpcRowDataGroupByShortName = make(map[string][]CpcRow)
	)
	for _, crow := range cpcRowDataList {
		company := crow.lessor
		companyGroup[company] = company
	}
	comStartNum := 0
	for _, snum := range companyStartColNumMap.(map[string]int) {
		if snum > comStartNum {
			comStartNum = snum
		}
	}
	for shorName, _ := range companyStartColNumMap.(map[string]int) {
		for longName, _ := range companyGroup {
			if strings.Contains(longName, shorName) {
				companyShortToLongName[longName] = shorName
			} else {
				_, ok := companyShortToLongName[longName]
				if longName != "" && !ok {
					companyShortToLongName[longName] = longName
					comStartNum = comStartNum + 3
					companyStartColNumMap.(map[string]int)[longName] = comStartNum
				}
			}
		}
	}

	for _, crow := range cpcRowDataList {
		longName := crow.lessor
		shotName, ok := companyShortToLongName[longName]
		if ok {
			cpcRowDataGroupByShortName[shotName] = append(cpcRowDataGroupByShortName[shotName], crow)
		}
	}
	var rowList []WycRow
	for shorName, datas := range cpcRowDataGroupByShortName {
		var cpcRowDataGroupByDate = make(map[string][]CpcRow)
		for _, data := range datas {
			dateStr := data.date
			cpcRowDataGroupByDate[dateStr] = append(cpcRowDataGroupByDate[dateStr], data)
		}
		for dateStr, datas1 := range cpcRowDataGroupByDate {
			var cpcRowDataGroupByFeeStr = make(map[string][]CpcRow)
			for _, data := range datas1 {
				feeStr := data.term + "-" + data.charge
				data.lessor = shorName
				cpcRowDataGroupByFeeStr[feeStr] = append(cpcRowDataGroupByFeeStr[feeStr], data)
			}
			for _, datas2 := range cpcRowDataGroupByFeeStr {
				var rowData WycRow
				for inx, data := range datas2 {
					rowData.fee = data.charge
					rowData.qishu = data.term
					rowData.taishu = inx + 1
					rowData.qijian = dateStr
					rowData.company = shorName
				}
				rowList = append(rowList, rowData)
			}
		}
	}

	var wycDataGroupByDateStr = make(map[string][]WycRow)
	for _, data := range rowList {
		dateStr := data.qijian
		wycDataGroupByDateStr[dateStr] = append(wycDataGroupByDateStr[dateStr], data)
	}
	var dateArr []string
	var dateQishuNum = make(map[string][]map[string]int)
	for dateStr, datas := range wycDataGroupByDateStr {
		var wycDtaGroupByQishu = make(map[string][]WycRow)
		for _, data := range datas {
			qishu := data.qishu
			wycDtaGroupByQishu[qishu] = append(wycDtaGroupByQishu[qishu], data)
		}
		var qishuList []map[string]int
		for qishustr, datas1 := range wycDtaGroupByQishu {
			var companyQishu = make(map[string]int)
			for _, data2 := range datas1 {
				num, ok := companyQishu[data2.company]
				if ok {
					companyQishu[data2.company] = num + 1
				} else {
					companyQishu[data2.company] = 1
				}
			}
			qishuNum := 0
			for _, num := range companyQishu {
				if num > qishuNum {
					qishuNum = num
				}
			}
			var qishustrnum = map[string]int{qishustr: qishuNum}
			qishuList = append(qishuList, qishustrnum)
		}
		dateQishuNum[dateStr] = qishuList
		dateArr = append(dateArr, dateStr)
	}
	sort.Strings(dateArr)

	startRowNum := wycfwfHeader["startRowNum"].(int) + 3
	sheetName := "网约车运营服务费"
	colNumStr := strconv.Itoa(startRowNum)
	for comstr, startnum := range companyStartColNumMap.(map[string]int) {
		s := NumToChar(startnum)
		e := NumToChar(startnum + 2)
		//fmt.Println(s + colNumStr + "---" + e + colNumStr + "---" + comstr)
		err := file.MergeCell(sheetName, s+colNumStr, e+colNumStr)
		if err != nil {
			panic(err)
		}
		file.SetCellValue(sheetName, s+colNumStr, comstr)
	}
	startRowNum = startRowNum + 1
	for _, dateStr := range dateArr {
		wycDataGroup := wycDataGroupByDateStr[dateStr]
		var wycDtaGroupByCompany = make(map[string][]WycRow)
		for _, cdata := range wycDataGroup {
			company := cdata.company
			wycDtaGroupByCompany[company] = append(wycDtaGroupByCompany[company], cdata)
		}
		qishuList := dateQishuNum[dateStr]
		for _, qishuNumMap := range qishuList {
			for qishustr, qishuNum := range qishuNumMap {
				var wycDtaGroupByCompanyQishu = make(map[string][]WycRow)
				for comStr, companyDatas := range wycDtaGroupByCompany {
					for _, comcompanyData := range companyDatas {
						if comcompanyData.qishu == qishustr {
							wycDtaGroupByCompanyQishu[comStr] = append(wycDtaGroupByCompanyQishu[comStr], comcompanyData)
						}
					}
				}
				for i := 0; i < qishuNum; i++ {
					startRowNum++
					colNumStr := strconv.Itoa(startRowNum)
					file.SetCellValue(sheetName, "A"+colNumStr, dateStr)
					file.SetCellValue(sheetName, "B"+colNumStr, qishustr)
					for comStr, companyDatas := range wycDtaGroupByCompanyQishu {
						comStart := companyStartColNumMap.(map[string]int)[comStr]
						if len(companyDatas) > i {
							companyData := companyDatas[i]
							file.SetCellValue(sheetName, NumToChar(comStart)+colNumStr, companyData.taishu)
							file.SetCellValue(sheetName, NumToChar(comStart+1)+colNumStr, companyData.fee)
						}
					}

				}
			}

		}

	}
	ZjceTongji(file, startRowNum)
	file.Save()
}

// 租金差额统计
func ZjceTongji(file *excelize.File, startRowNum int) {
	sheetName1 := "网约车B端对账单"
	rows, err := file.GetRows(sheetName1)
	if err != nil {
		panic(err)
	}
	var dateStrMap = make(map[string]string)
	var dateToZjMap = make(map[string][]string)
	for inx, row := range rows {
		if inx == 0 {
			continue
		}
		zj := row[12]
		zjNum, err := strconv.Atoi(zj)
		if err == nil {
			if zjNum > 0 {
				colStr := row[0]
				dateArr := strings.Split(colStr, "/")
				if len(dateArr) == 3 {
					dateStr := dateArr[0] + "年" + dateArr[1] + "月"
					dateToZjMap[dateStr] = append(dateToZjMap[dateStr], zj)
					dateStrMap[dateStr] = dateStr
				} else {
					continue
				}
			}
		}
	}

	sheetName2 := "网约车花芝租租金"
	rows2, err := file.GetRows(sheetName2)
	if err != nil {
		panic(err)
	}
	var dateToZjMap2 = make(map[string][]string)
	for inx, row := range rows2 {
		if inx == 0 {
			continue
		}
		zj := row[28]
		zjNum, err := strconv.Atoi(zj)
		if err == nil {
			if zjNum > 0 {
				colStr := row[0]
				dateArr := strings.Split(colStr, "/")
				if len(dateArr) == 3 {
					dateStr := dateArr[0] + "年" + dateArr[1] + "月"
					dateToZjMap2[dateStr] = append(dateToZjMap2[dateStr], zj)
					dateStrMap[dateStr] = dateStr
				} else {
					continue
				}
			}
		}
	}
	var dateArr []string
	for _, datestr := range dateStrMap {
		dateArr = append(dateArr, datestr)
	}
	sort.Strings(dateArr)
	dateToZjMapList := make(map[string][]string)
	for _, dateStr := range dateArr {
		arr, ok := dateToZjMap[dateStr]
		if ok {
			dateToZjMapList[dateStr] = append(dateToZjMapList[dateStr], arr...)
		}
		arr1, ok := dateToZjMap2[dateStr]
		if ok {
			dateToZjMapList[dateStr] = append(dateToZjMapList[dateStr], arr1...)
		}
	}

	rowNum := startRowNum + 3
	sheetName := "网约车运营服务费"
	file.SetCellValue(sheetName, "B"+strconv.Itoa(rowNum), "租金差")
	file.SetCellValue(sheetName, "C"+strconv.Itoa(rowNum), "数量")
	for dateStr, arr := range dateToZjMapList {
		var zjNums = make(map[string]int)
		for _, zj := range arr {
			num, ok := zjNums[zj]
			if ok {
				zjNums[zj] = num + 1
			} else {
				zjNums[zj] = 1
			}
		}

		for zj, zjNum := range zjNums {
			rowNum++
			rowNumStr := strconv.Itoa(rowNum)
			file.SetCellValue(sheetName, "A"+rowNumStr, dateStr)
			file.SetCellValue(sheetName, "B"+rowNumStr, zj)
			file.SetCellValue(sheetName, "C"+rowNumStr, zjNum)
		}
	}

}

func NumToChar(num int) string {
	tostr := ""
	anum := num + 1
	if anum < 27 {
		tostr = string('A' + (anum - 1))
	} else {
		an := math.Floor(float64(anum) / 26)
		astr := string('A' + (int(an) - 1))
		bn := math.Mod(float64(anum), 26)
		bstr := string('A' + (int(bn) - 1))
		tostr = astr + bstr
	}
	return tostr
}
