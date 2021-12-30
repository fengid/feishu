package feishu

import (
	"bytes"
	"github.com/json-iterator/go"
	"log"
	"net/http"
	"net/url"
	"strconv"
)

/**************************************************** 飞书云文档api ****************************************************/

// 文档类型
const (
	DOC      = "doc"      // 飞书文档
	SHEET    = "sheet"    // 飞书电子表格
	BITABLE  = "bitable"  // 飞书多维表格
	MINDNOTE = "mindnote" // 飞书思维笔记
	FILE     = "file"     // 飞书文件
)

type Response struct {
	Code int    `json:"code"`
	Msg  string `json:"msg"`
}

/************************* 操作文件夹 ******************************/

type GetRootFolderTokenResponse struct {
	Response
	Data struct {
		Id     string `json:"id"`
		Token  string `json:"token"`
		UserId string `json:"user_id"`
	} `json:"data"`
}

// GetRootFolderToken 获取根目录的token
func (c *Client) GetRootFolderToken() (*GetRootFolderTokenResponse, error) {
	request, _ := http.NewRequest(http.MethodGet, ServerUrl+"/open-apis/drive/explorer/v2/root_folder/meta", nil)
	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &GetRootFolderTokenResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

type GetFolderChildrenRequest struct {
	FolderToken string   `json:"-"`
	Types       []string `json:"-"`
}

type GetFolderChildrenResponse struct {
	Data struct {
		ParentToken string                  `json:"parentToken"`
		Children    map[string]ChildrenInfo `json:"children"`
	} `json:"data"`
}

type ChildrenInfo struct {
	Token string `json:"token"`
	Name  string `json:"name"`
	Type  string `json:"type"`
}

// GetFolderChildren 获取目录下所有文档信息
func (c *Client) GetFolderChildren(args *GetFolderChildrenRequest)(*GetFolderChildrenResponse, error) {
	param := url.Values{}
	for _, v := range args.Types {
		param.Add("types", v)
	}
	request, _ := http.NewRequest(http.MethodGet, ServerUrl+"/open-apis/drive/explorer/v2/folder/"+
		args.FolderToken+"/children?"+param.Encode(), nil)
	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &GetFolderChildrenResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

/************************* 操作文档 ******************************/

type GetFileInfoRequest struct {
	RequestDocs []*DocsInfo `json:"request_docs"`
}

type DocsInfo struct {
	DocsToken string `json:"docs_token"`
	DocsType  string `json:"docs_type"`
}

type GetFileInfoResponse struct {
	Response
	Data struct {
		DocsMetas []struct {
			CreateTime       int    `json:"create_time"`
			DocsToken        string `json:"docs_token"`
			DocsType         string `json:"docs_type"`
			LatestModifyTime int    `json:"latest_modify_time"`
			LatestModifyUser string `json:"latest_modify_user"`
			OwnerId          string `json:"owner_id"`
			Title            string `json:"title"`
		} `json:"docs_metas"`
	} `json:"data"`
}

// GetFileInfo 获取文档源信息
func (c *Client) GetFileInfo(args *GetFileInfoRequest) (*GetFileInfoResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/suite/docs-api/meta", buff)
	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &GetFileInfoResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

/************************* 操作权限 ******************************/

// 用户类型
const (
	EMAIL              = "email"            // 飞书企业邮箱
	OPEN_ID            = "openid"           // 开放平台ID
	OPEN_CHAT          = "openchat"         // 开放平台群组
	OPEN_DEPARTMENT_ID = "opendepartmentid" // 开放平台部门ID
	USER_ID            = "userid"           // 用户自定义ID（支持应用身份）
)

// 权限类型
const (
	VIEW        = "view"        // 可阅读
	EDIT        = "edit"        // 可编辑
	FULL_ACCESS = "full_access" // 所有权限
)

type AddPermissionRequest struct {
	Type             string `json:"-"`
	NeedNotification bool   `json:"-"`
	MemberType       string `json:"member_type"`
	MemberId         string `json:"member_id"`
	Perm             string `json:"perm"`
	FileToken        string `json:"-"`
}

type AddPermissionResponse struct {
	Data struct {
		Member struct {
			MemberId   string `json:"member_id"`
			MemberType string `json:"member_type"`
			Perm       string `json:"perm"`
		} `json:"member"`
	} `json:"data"`
}

// AddPermission 给文档添加协作者
func (c *Client) AddPermission(args *AddPermissionRequest) (*AddPermissionResponse, error) {
	params := url.Values{}
	params.Add("type", args.Type)
	params.Add("need_notification", strconv.FormatBool(args.NeedNotification))

	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/drive/v1/permissions/"+args.FileToken+
		"/members?"+params.Encode(), buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &AddPermissionResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

/************************* 电子表格 ******************************/

type CreateSpreadsheetRequest struct {
	Title       string `json:"title"`
	FolderToken string `json:"folder_token"`
}

type CreateSpreadsheetResponse struct {
	Response
	Data struct {
		Spreadsheet struct {
			Title            string `json:"title"`
			FolderToken      string `json:"folder_token"`
			Url              string `json:"url"`
			SpreadsheetToken string `json:"spreadsheet_token"`
		} `json:"spreadsheet"`
	} `json:"data"`
}

// CreateSpreadsheet 创建电子表格
func (c *Client) CreateSpreadsheet(args *CreateSpreadsheetRequest) (*CreateSpreadsheetResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/sheets/v3/spreadsheets", buff)
	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &CreateSpreadsheetResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

/*************** 工作表sheet *************************/

// MexSheet 最大sheet数
const MexSheet = 999999999

type GetSheetInfoRequest struct {
	SpreadsheetToken string
	ExtFields        string
	UserIdType       string
}

type GetSheetInfoResponse struct {
	Response
	Data struct {
		Properties struct {
			Title      string `json:"title"`
			OwnerUser  int    `json:"ownerUser"`
			SheetCount int    `json:"sheetCount"`
			Revision   int    `json:"revision"`
		} `json:"properties"`
		Sheets []struct {
			SheetId        string `json:"sheetId"`
			Title          string `json:"title"`
			Index          int    `json:"index"`
			RowCount       int    `json:"rowCount"`
			ColumnCount    int    `json:"columnCount"`
			FrozenColCount int    `json:"frozenColCount"`
			FrozenRowCount int    `json:"frozenRowCount"`
			Merges         []struct {
				ColumnCount      int `json:"columnCount"`
				RowCount         int `json:"rowCount"`
				StartColumnIndex int `json:"startColumnIndex"`
				StartRowIndex    int `json:"startRowIndex"`
			} `json:"merges,omitempty"`
			ProtectedRange []struct {
				Dimension struct {
					EndIndex       int    `json:"endIndex"`
					MajorDimension string `json:"majorDimension"`
					SheetId        string `json:"sheetId"`
					StartIndex     int    `json:"startIndex"`
				} `json:"dimension"`
				ProtectId string `json:"protectId"`
				SheetId   string `json:"sheetId"`
				LockInfo  string `json:"lockInfo"`
			} `json:"protectedRange,omitempty"`
			BlockInfo struct {
				BlockToken string `json:"blockToken"`
				BlockType  string `json:"blockType"`
			} `json:"blockInfo,omitempty"`
		} `json:"sheets"`
		SpreadsheetToken string `json:"spreadsheetToken"`
	} `json:"data"`
}

// GetSheetInfo 获取工作表信息
func (c *Client) GetSheetInfo(args *GetSheetInfoRequest) (*GetSheetInfoResponse, error) {
	param := url.Values{}
	param.Add("extFields", args.ExtFields)
	param.Add("user_id_type", args.UserIdType)

	request, _ := http.NewRequest(http.MethodGet, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/metainfo?"+param.Encode(), nil)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &GetSheetInfoResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

type OperationSheetRequest struct {
	SpreadsheetToken string            `json:"-"`
	Requests         []OperationObject `json:"requests"` // AddSheet CopySheet DeleteSheet
}

type OperationObject struct {
	AddSheet    *AddSheet    `json:"addSheet,omitempty"`
	CopySheet   *CopySheet   `json:"copySheet,omitempty"`
	DeleteSheet *DeleteSheet `json:"deleteSheet,omitempty"`
}

// AddSheet 添加sheet
type AddSheet struct {
	Properties Properties `json:"properties"`
}

type Properties struct {
	Title string `json:"title"`
	Index int    `json:"index"`
}

// CopySheet 复制sheet
type CopySheet struct {
	Source      Source      `json:"source"`
	Destination Destination `json:"destination"`
}

type Source struct {
	SheetId string `json:"sheetId"`
}

type Destination struct {
	Title string `json:"title"`
}

// DeleteSheet 删除sheet
type DeleteSheet struct {
	SheetId string `json:"sheetId"`
}

type OperationSheetResponse struct {
	Response
	Data struct {
		Replies []OperationObjectResponse `json:"replies"`
	} `json:"data"`
}

type OperationObjectResponse struct {
	AddSheet struct {
		Properties struct {
			SheetId string `json:"sheetId"`
			Title   string `json:"title"`
			Index   int    `json:"index"`
		} `json:"properties"`
	} `json:"addSheet,omitempty"`
	CopySheet struct {
		Properties struct {
			SheetId string `json:"sheetId"`
			Title   string `json:"title"`
			Index   int    `json:"index"`
		} `json:"properties"`
	} `json:"copySheet,omitempty"`
	DeleteSheet struct {
		Result  bool   `json:"result"`
		SheetId string `json:"sheetId"`
	} `json:"deleteSheet,omitempty"`
}

// OperationSheet 操作工作表 (增加, 复制, 删除)
func (c *Client) OperationSheet(args *OperationSheetRequest) (*OperationSheetResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	log.Println(string(body))
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/sheets_batch_update", buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &OperationSheetResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

/***************** 行列数据操作 *************************/

type ReadRangeRequest struct {
	SpreadsheetToken     string
	Range                string
	ValueRenderOption    string
	DateTimeRenderOption string
	UserIdType           string
}

type ReadRangeResponse struct {
	Response
	Data struct {
		Revision         int    `json:"revision"`
		SpreadsheetToken string `json:"spreadsheetToken"`
		ValueRange       struct {
			MajorDimension string          `json:"majorDimension"`
			Range          string          `json:"range"`
			Revision       int             `json:"revision"`
			Values         [][]interface{} `json:"values"`
		} `json:"valueRange"`
	} `json:"data"`
}

// ReadRange 读取单个范围数据
func (c *Client) ReadRange(args *ReadRangeRequest) (*ReadRangeResponse, error) {
	param := url.Values{}
	param.Add("valueRenderOption", args.ValueRenderOption)
	param.Add("dateTimeRenderOption", args.DateTimeRenderOption)
	param.Add("user_id_type", args.UserIdType)

	request, _ := http.NewRequest(http.MethodGet, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/values/"+args.Range+"?"+param.Encode(), nil)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &ReadRangeResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}

	return response, nil
}

type WriteRangeRequest struct {
	SpreadsheetToken string     `json:"-"`
	ValueRange       ValueRange `json:"valueRange"`
}

type ValueRange struct {
	Range  string          `json:"range"`
	Values [][]interface{} `json:"values"`
}

type WriteRangeResponse struct {
	Response
	Data struct {
		Revision         int    `json:"revision"`
		SpreadsheetToken string `json:"spreadsheetToken"`
		UpdatedCells     int    `json:"updatedCells"`
		UpdatedColumns   int    `json:"updatedColumns"`
		UpdatedRange     string `json:"updatedRange"`
		UpdatedRows      int    `json:"updatedRows"`
	} `json:"data"`
}

// WriteRange 单个范围写入数据
func (c *Client) WriteRange(args *WriteRangeRequest) (*WriteRangeResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPut, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/values", buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &WriteRangeResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}
	return response, nil
}

type SetCellStyleRequest struct {
	SpreadsheetToken string      `json:"-"`
	AppendStyle      AppendStyle `json:"appendStyle"`
}

type AppendStyle struct {
	Range string `json:"range"`
	Style Style  `json:"style"`
}

type Style struct {
	Font           *Font   `json:"font,omitempty"`
	TextDecoration *int    `json:"textDecoration,omitempty"` // 文本装饰 ，0 默认，1 下划线，2 删除线 ，3 下划线和删除线
	Formatter      *string `json:"formatter,omitempty"`
	HAlign         *int    `json:"hAlign,omitempty"`    // 水平对齐，0 左对齐，1 中对齐，2 右对齐
	VAlign         *int    `json:"vAlign,omitempty"`    // 垂直对齐， 0 上对齐，1 中对齐， 2 下对齐
	ForeColor      *string `json:"foreColor,omitempty"` // 字体颜色
	BackColor      *string `json:"backColor,omitempty"` // 背景颜色
	// "FULL_BORDER"，"OUTER_BORDER"，"INNER_BORDER"，"NO_BORDER"，"LEFT_BORDER"，"RIGHT_BORDER"，"TOP_BORDER"，"BOTTOM_BORDER"
	BorderType  *string `json:"borderType,omitempty"`
	BorderColor *string `json:"borderColor,omitempty"` // 边框颜色
	Clean       *bool   `json:"clean,omitempty"`
}

type Font struct {
	Bold     bool   `json:"bold"`     // 是否加粗
	Italic   bool   `json:"italic"`   // 是否斜体
	FontSize string `json:"fontSize"` // 字体大小 字号大小为9~36 行距固定为1.5，如:10pt/1.5
	Clean    bool   `json:"clean"`    // 清除 font 格式,默认 false
}

type SetCellStyleResponse struct {
	Response
	Data struct {
		SpreadsheetToken string `json:"spreadsheetToken"`
		UpdatedRange     string `json:"updatedRange"`
		UpdatedRows      int    `json:"updatedRows"`
		UpdatedColumns   int    `json:"updatedColumns"`
		UpdatedCells     int    `json:"updatedCells"`
		Revision         int    `json:"revision"`
	} `json:"data"`
}

// SetCellStyle 设置单元格样式
func (c *Client) SetCellStyle(args *SetCellStyleRequest) (*SetCellStyleResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPut, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/style", buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &SetCellStyleResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}
	return response, nil
}

type SetCellStyleBatchData struct {
	Range []string `json:"ranges"`
	Style Style    `json:"style"`
}

type SetCellStyleBatchRequest struct {
	SpreadsheetToken string                  `json:"-"`
	Data             []SetCellStyleBatchData `json:"data"`
}

type SetCellStyleBatchResponse struct {
	Response
	Data struct {
		SpreadsheetToken    string `json:"spreadsheetToken"`
		TotalUpdatedCells   int    `json:"totalUpdatedCells"`
		TotalUpdatedColumns int    `json:"totalUpdatedColumns"`
		TotalUpdatedRows    int    `json:"totalUpdatedRows"`
		Revision            int    `json:"revision"`
		Responses           []struct {
			SpreadsheetToken string `json:"spreadsheetToken"`
			UpdatedRange     string `json:"updatedRange"`
			UpdatedRows      int    `json:"updatedRows"`
			UpdatedColumns   int    `json:"updatedColumns"`
			UpdatedCells     int    `json:"updatedCells"`
		} `json:"responses"`
	} `json:"data"`
}

// SetCellStyleBatch 批量设置样式
func (c *Client) SetCellStyleBatch(args *SetCellStyleBatchRequest) (*SetCellStyleBatchResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPut, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/styles_batch_update", buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &SetCellStyleBatchResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}
	return response, nil
}

type FindRequest struct {
	SpreadsheetToken string        `json:"-"`
	SheetId          string        `json:"-"`
	FindCondition    FindCondition `json:"find_condition"`
	Find             string        `json:"find"`
}

type FindCondition struct {
	Range           string `json:"range"`
	MatchCase       bool   `json:"match_case"`
	MatchEntireCell bool   `json:"match_entire_cell"`
	SearchByRegex   bool   `json:"search_by_regex"`
	IncludeFormulas bool   `json:"include_formulas"`
}

type FindResponse struct {
	Response
	Data struct {
		FindResult struct {
			MatchedCells        []string `json:"matched_cells"`
			MatchedFormulaCells []string `json:"matched_formula_cells"`
			RowsCount           int      `json:"rows_count"`
		} `json:"find_result"`
	} `json:"data"`
}

// Find 查找
func (c *Client) Find(args *FindRequest) (*FindResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/sheets/v3/spreadsheets/"+
		args.SpreadsheetToken+"/sheets/"+args.SheetId+"/find", buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &FindResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}
	return response, nil
}

const (
	ROWS    = "ROWS"    // 行
	COLUMNS = "COLUMNS" // 列

	BEFORE = "BEFORE"
	AFTER  = "AFTER"
)

type InsertDimensionRangeRequest struct {
	SpreadsheetToken string    `json:"-"`
	Dimension        Dimension `json:"dimension"`
	InheritStyle     string    `json:"inheritStyle"`
}

type Dimension struct {
	SheetId        string `json:"sheetId"`
	MajorDimension string `json:"majorDimension"`
	StartIndex     int    `json:"startIndex"`
	EndIndex       int    `json:"endIndex"`
}

type InsertDimensionRangeResponse struct {
	Response
	Data struct{} `json:"data"`
}

// InsertDimensionRange 插入行列
func (c *Client) InsertDimensionRange(args *InsertDimensionRangeRequest) (*InsertDimensionRangeResponse, error) {
	body, err := jsoniter.Marshal(args)
	if err != nil {
		return nil, err
	}
	buff := bytes.NewBuffer(body)
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/sheets/v2/spreadsheets/"+
		args.SpreadsheetToken+"/insert_dimension_range", buff)

	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}

	response := &InsertDimensionRangeResponse{}
	if err = jsoniter.Unmarshal(resp, response); err != nil {
		return nil, err
	}
	return response, nil
}
