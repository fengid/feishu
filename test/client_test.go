package test

import (
	"fmt"
	"github.com/fengid/feishu"
	"log"
	"testing"
	"time"
)

const (
	AppId = "cli_a14ef7fc7f781013"
	AppSecret = "0t9w2D4OyDGonV7g7gckabq8wOI0Q48J"
)

var Client = feishu.NewClient(AppId, AppSecret)

// 获取根目录信息
func TestGetRootFolderToken(t *testing.T) {
	res, err := Client.GetRootFolderToken()
	if err != nil {
		panic(err)
	}

	log.Println(res.Data)
}

// 获取文档源信息
func TestGetFileInfo(t *testing.T) {
	requestDocs :=	&feishu.DocsInfo{
		DocsToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		DocsType: "sheet",
	}
	args := &feishu.GetFileInfoRequest{RequestDocs: []*feishu.DocsInfo{requestDocs}}
	res, err := Client.GetFileInfo(args)
	if err != nil {
		panic(err)
	}

	log.Println(res.Data)
}

// 给文档添加协作者
func TestAddPermission(t *testing.T) {
	args := &feishu.AddPermissionRequest{
		Type: feishu.SHEET,
		NeedNotification: true,
		MemberType: feishu.OPEN_ID,
		MemberId: "ou_0ddcdd365511f763385a916757bc7f93",
		Perm: feishu.EDIT,
		FileToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
	}
	res, err := Client.AddPermission(args)
	if err != nil {
		panic(err)
	}

	log.Println(res.Data)
}

// 创建电子表格
func TestCreateSpreadsheet(t *testing.T) {
	args := &feishu.CreateSpreadsheetRequest{
		Title: "测试6",
		FolderToken: "nodcn3JpjzMZZ0ngo2bTdWKE6wc",
	}
	res, err := Client.CreateSpreadsheet(args)
	if err != nil {
		panic(err)
	}

	log.Println(res.Data)
}

// 获取工作表信息
func TestGetSheetInfo(t *testing.T) {
	args := &feishu.GetSheetInfoRequest{
		SpreadsheetToken: "shtcnylYjqdPEfyLpNXSpJeSjGc",
	}
	res, err := Client.GetSheetInfo(args)
	if err != nil {
		panic(err)
	}

	log.Println(res.Data)
}

// 操作工作表
func TestOperationSheet(t *testing.T) {
	add := &feishu.AddSheet{
		Properties: feishu.Properties{
			Title: "13月",
			Index: feishu.MexSheet,
		},
	}
	args := &feishu.OperationSheetRequest{
		SpreadsheetToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		Requests: []feishu.OperationObject{
			{AddSheet: add},
		},
	}
	res, err := Client.OperationSheet(args)
	if err != nil {
		panic(err)
	}

	log.Println(res)
}

// 读取某个范围的数据
func TestReadRange(t *testing.T) {
	args := &feishu.ReadRangeRequest{
		SpreadsheetToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		Range: "3Kk4Fy",
	}
	res, err := Client.ReadRange(args)
	if err != nil {
		panic(err)
	}

	log.Println(res)
}

// 设置样式
func TestSetCellStyleRequest(t *testing.T) {
	color := "#E4EFDC"
	args := &feishu.SetCellStyleRequest{
		SpreadsheetToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		AppendStyle: feishu.AppendStyle{
			Range: "3Kk4Fy!A:B",
			Style: feishu.Style{
				BackColor: &color,
			},
		},
	}
	res, err := Client.SetCellStyle(args)
	if err != nil {
		panic(err)
	}

	log.Println(res)
}

// 批量设置样式
func TestSetCellStyleBatchRequest(t *testing.T) {
	color := "#E4EFDC"
	args := &feishu.SetCellStyleBatchRequest{
		SpreadsheetToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		Data: []feishu.SetCellStyleBatchData{
			{
				Range: []string{"3Kk4Fy!A:B"},
				Style: feishu.Style{
					BackColor: &color,
				},
			},
		},
	}
	res, err := Client.SetCellStyleBatch(args)
	if err != nil {
		panic(err)
	}

	log.Println(res)
}

// 查询结果
func TestFind(t *testing.T) {
	args := &feishu.FindRequest{
		SpreadsheetToken: "shtcnvwqY6WkMRAbbhB98LqD0Bg",
		SheetId: "2xGtW",
		FindCondition: feishu.FindCondition{
			Range: "2xGtW",
			MatchCase: true,
		},
		Find: "第1场",
	}
	res, err := Client.Find(args)
	if err != nil {
		panic(err)
	}

	log.Println(res.Data.FindResult.MatchedCells)
}

// 插入行列
func TestInsertDimensionRange(t *testing.T) {
	args := &feishu.InsertDimensionRangeRequest{
		SpreadsheetToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		Dimension: feishu.Dimension{
			SheetId: "3Kk4Fy",
			MajorDimension: feishu.ROWS,
			StartIndex: 4,
			EndIndex: 5,
		},
		InheritStyle: feishu.BEFORE,
	}
	res, err := Client.InsertDimensionRange(args)
	if err != nil {
		panic(err)
	}
	log.Println(Client.TokenManager.GetAccessToken())
	log.Println(res)
}

func GenWeekAndDay(year, month int)([]interface{}, []interface{}){
	monthDays := 0
	totalDyas := 0
	isLeapYear := false

	// 计算是不是润年
	if year%400 == 0 || (year%4 == 0 && year%100 != 0) {
		isLeapYear = true
	}

	//计算距离1900年的天数
	for i := 1900; i < year; i++ {
		if i%400 == 0 || (i%4 == 0 && i%100 != 0) {
			totalDyas += 366
		} else {
			totalDyas += 365
		}
	}

	//计算截至到当月1号前的总天数
	for j := 1; j <= month; j++ {
		switch j {
		case 1, 3, 5, 7, 8, 10, 12:
			monthDays = 31
			break
		case 2:
			if isLeapYear {
				monthDays = 29
			} else {
				monthDays = 28
			}
			break
		case 4, 6, 9, 11:
			monthDays = 30
			break
		default:
			fmt.Println("input month is error.")
		}

		if j != month {
			totalDyas += monthDays
		}
	}

	//计算当月1号是星期几 0为周日
	weekDay := 0
	weekDay = totalDyas%7
	if weekDay == 7 {
		weekDay = 0
	}

	day := make([]interface{}, 0, monthDays)
	week := make([]interface{}, 0, monthDays)
	day = append(day, "日期")
	week = append(week, "星期")
	for m:=1; m<=monthDays; m++{
		day = append(day, fmt.Sprintf("%d月%d号", month, m))
		weekInt := (weekDay+m)%7
		week = append(week, weeks[weekInt])
	}

	return day, week
}

var weeks = map[int]string{
	0: "星期日",
	1: "星期一",
	2: "星期二",
	3: "星期三",
	4: "星期四",
	5: "星期五",
	6: "星期六",
}

// 测试生成初始化模板
func TestModel(t *testing.T) {
	day, week := GenWeekAndDay(2021, 12)

	args := &feishu.WriteRangeRequest{
		SpreadsheetToken: "shtcnvPpGKEHsJFst0L7bQwu9pd",
		ValueRange: feishu.ValueRange{
			Range: "3Kk4Fy",
			Values: [][]interface{}{
				day,
				week,
				{"开播时间段"},
				{"时长"},
			},
		},
	}
	res, err := Client.WriteRange(args)
	if err != nil {
		panic(err)
	}

	log.Println(res)
}

// 目录下所有文档信息
func TestGetFolderChildren(t *testing.T) {
	args := &feishu.GetFolderChildrenRequest{
		FolderToken: "nodcn3JpjzMZZ0ngo2bTdWKE6wc",
		Types: []string{feishu.SHEET},
	}
	res, err := Client.GetFolderChildren(args)
	if err != nil {
		panic(err)
	}

	log.Println(res)

	for _, child := range res.Data.Children{
		_, err = Client.DeleteSheet(child.Token)
		if err != nil{
			log.Println(err)
		}
		log.Println(child.Token)
		time.Sleep(time.Second)
	}
}