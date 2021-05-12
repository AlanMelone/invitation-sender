package invitation_sender

import (
	"errors"
	"fmt"
	"net/smtp"
	"os"

	"code.sajari.com/docconv"
	"github.com/nguyenthenguyen/docx"
	"github.com/tealeg/xlsx"
)

func SendEmailFromTemplate(tablePath string, sheetName string, templatePath string) (string, error) {
	wb, err := xlsx.OpenFile(tablePath)
	if err != nil {
		panic(err)
	}

	fmt.Println("Sheets in this file:")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("----")

	sh, ok := wb.Sheet[sheetName]
	if !ok {
		return "", errors.New("Sheet doesn't exist")
	}
	fmt.Println("Max row in sheet:", sh.MaxRow)

	keyMap := make(map[string]int)

	arrayLimit := 4

	parametersMap := make(map[string]string)

	parametersRow := sh.Row(0)
	for index, cell := range parametersRow.Cells {
		if cell.Value != "" {
			keyMap[cell.Value] = index
		}
	}
	fmt.Println("Key Map:")
	fmt.Println(keyMap)
	for i, row := range sh.Rows {
		if i == 0 {
			continue
		}
		for index, cell := range row.Cells {
			for k, v := range keyMap {
				if index == v {
					parametersMap[k] = cell.Value
				}
			}
		}
		if parametersMap["Email"] == "" {
			continue
		}

		text, err := getChangedDocumentContent(templatePath, parametersMap)

		if err != nil {
			return "", err
		}

		if err := sendMail(parametersMap["Email"], text); err != nil {
			return "", err
		}

		fmt.Println(text)
		if i == arrayLimit {
			break
		}
	}

	return "", nil
}

func getChangedDocumentContent(templatePath string, parametersMap map[string]string) (string, error) {
	r, err := docx.ReadDocxFile(templatePath)
	if err != nil {
		return "", err
	}
	defer r.Close()

	doc := r.Editable()

	fmt.Println("Old Content")
	fmt.Println(doc.GetContent())

	for k, v := range parametersMap {
		fmt.Println("Key: " + k)
		fmt.Println("Value: " + v)
		fmt.Println("---")
		err := doc.Replace("!"+k, v, -1)
		if err != nil {
			return "", err
		}
	}

	doc.WriteToFile("temp.docx")

	res, err := docconv.ConvertPath("temp.docx")
	return res.Body, nil
}

func sendMail(to string, message string) error {
	from := os.Getenv("MAIL_FROM")
	password := os.Getenv("MAIL_PASSWD")

	toList := []string{to}

	if password == "" {
		return errors.New("Enter MAIL_PASSWD variable to env")
	}

	if from == "" {
		return errors.New("Enter MAIL_FROM variable to env")
	}

	host := "mail.nic.ru"
	port := "25"

	body := []byte(
		"Subject: IBCM Invitation\r\n" +
			"\r\n" + message)

	auth := smtp.PlainAuth("", from, password, host)
	err := smtp.SendMail(host+":"+port, auth, from, toList, body)
	if err != nil {
		return err
	}
	return nil
}
