package feishu

import (
	"encoding/json"
	"net/http"
	"strings"
)

type InteractiveV1CardUpdateParam struct {
	Token string                           `json:"token"`
	Card  InteractiveV1CardUpdateCardParam `json:"card"`
}

type InteractiveV1CardUpdateCardParam struct {
	OpenIds  []string `json:"open_ids"`
	Config   interface{}
	Header   interface{}
	Elements interface{}
}

type InteractiveV1CardUpdateRes struct {
	ResponseCode
}

// InteractiveV1CardUpdate 消息卡片延迟更新
func (c *Client) InteractiveV1CardUpdate(param InteractiveV1CardUpdateParam) (*InteractiveV1CardUpdateRes, error) {
	paramByte, err := json.Marshal(param)
	if err != nil {
		return nil, err
	}
	request, _ := http.NewRequest(http.MethodPost, ServerUrl+"/open-apis/interactive/v1/card/update", strings.NewReader(string(paramByte)))
	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}
	var data InteractiveV1CardUpdateRes
	err = json.Unmarshal(resp, &data)
	if err != nil {
		return nil, err
	}
	return &data, err
}

type MessageParam struct {
	MessageId string
}

type MessageRes struct {
	ResponseCode
	Data MessageResData `json:"data"`
}

type MessageResData struct {
	Items []MessageResDataItem `json:"items"`
}

type MessageResDataItem struct {
	MessageId      string                       `json:"message_id"`
	RootId         string                       `json:"root_id"`
	ParentId       string                       `json:"parent_id"`
	MsgType        string                       `json:"msg_type"`
	CreateTime     string                       `json:"create_time"`
	UpdateTime     string                       `json:"update_time"`
	Deleted        bool                         `json:"deleted"`
	Updated        bool                         `json:"updated"`
	ChatId         string                       `json:"chat_id"`
	Sender         MessageResDataItemSender     `json:"sender"`
	Body           MessageResDataItemBody       `json:"body"`
	Mentions       []MessageResDataItemMentions `json:"mentions"`
	UpperMessageId string                       `json:"upper_message_id"`
}

type MessageResDataItemSender struct {
	Id         string `json:"id"`
	IdType     string `json:"id_type"`
	SenderType string `json:"sender_type"`
}

type MessageResDataItemBody struct {
	Content string `json:"content"`
}

type MessageResDataItemMentions struct {
	Key    string `json:"key"`
	Id     string `json:"id"`
	IdType string `json:"id_type"`
	Name   string `json:"name"`
}

// Messages 获取指定消息的内容
func (c *Client) Messages(param MessageParam) (*MessageRes, error) {
	request, _ := http.NewRequest(http.MethodGet, ServerUrl+"/open-apis/im/v1/messages/"+param.MessageId, nil)
	AccessToken, err := c.TokenManager.GetAccessToken()
	if err != nil {
		return nil, err
	}
	resp, err := c.Do(request, AccessToken)
	if err != nil {
		return nil, err
	}
	var data MessageRes
	err = json.Unmarshal(resp, &data)
	if err != nil {
		return nil, err
	}
	return &data, err
}
