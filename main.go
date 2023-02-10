package main

import (
	"context"
	"encoding/json"
	"fmt"
	"golang.org/x/net/html"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"io"
	"log"
	"net/http"
	"os"
	"strconv"
	"strings"
)

type Table struct {
	columns    []Column
	sheetRange string
	converted  [][]interface{}
}
type Column struct {
	name string
	row  []string
}

//this method add new column to table
func (tbl *Table) Add(clmn Column) {
	tbl.columns = append(tbl.columns, clmn)
}

//this method make range for exporting to google sheets
func (tbl *Table) FindRange(shtName string, insertColumnNames bool) {
	var tempRange string
	var addRow int
	if insertColumnNames {
		addRow = 1
	} else {
		addRow = 0
	}
	lRange := "A1"
	rRange := string(rune(len(tbl.columns)+64)) + strconv.Itoa(len(tbl.columns[0].row)+addRow)
	tempRange = shtName + "!" + lRange + ":" + rRange
	tbl.sheetRange = tempRange
}

//this method convert type Table to [][] interface for exporting to google sheets
func (tbl *Table) Convert(insertColumnNames bool) {
	var cnv [][]interface{}
	//var clm interface{}

	for _, c := range tbl.columns {
		row := make([]interface{}, len(c.row))
		for i, v := range c.row {
			row[i] = v
		}
		if insertColumnNames {
			row = append([]interface{}{c.name}, row...)
		}
		//clm = interface{}(c.name)
		cnv = append(cnv, row)
	}

	tbl.converted = cnv
}

// this func get html page as string
func getHtmlPage(webPage string) (string, error) {

	resp, err := http.Get(webPage)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", err
	}

	return string(body), nil
}

// this func parse html table and convert it to type Table
func parseHTMLTable(text string, tbl *Table) {
	tkn := html.NewTokenizer(strings.NewReader(text))
	var isRowValue bool
	var isColumnName bool
	var clmnCounter int
	var clmnQty int
	var tempStr string
	for {
		tt := tkn.Next()
		t := tkn.Token()
		switch {
		case tt == html.ErrorToken:
			return
		case tt == html.StartTagToken && (t.Data == "td" || t.Data == "th"):
			if string(tkn.Raw()) == "<td class=\"confluenceTd\">" {
				clmnQty = len(tbl.columns) - 1
				isRowValue = true
				tempStr = ""
			}
			if t.Data == "th" {
				isColumnName = true
			}
		case tt == html.TextToken:
			if isColumnName {
				clmn := new(Column)
				clmn.name = string(tkn.Raw())
				clmn.row = make([]string, 0)
				tbl.Add(*clmn)
			}
			if isRowValue {
				tempStr += string(tkn.Raw()) + " "
			}

		case tt == html.EndTagToken && (t.Data == "td" || t.Data == "th"):
			if isRowValue {
				tempStr = strings.ReplaceAll(tempStr, "  ", " ")
				strings.TrimSuffix(tempStr, " ")
				tbl.columns[clmnCounter].row = append(tbl.columns[clmnCounter].row, tempStr)
				clmnCounter++
				if clmnQty < clmnCounter {
					clmnCounter = 0
				}
			}
			isRowValue = false
			isColumnName = false
		}
	}
}

// functons below - for OAuth2 to GoogleAPI

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

func main() {
	webPage := "https://confluence.hflabs.ru/pages/viewpage.action?pageId=1181220999"
	data, err := getHtmlPage(webPage)
	if err != nil {
		log.Fatal(err)
		return
	}
	//init new Table
	tbl := new(Table)
	//parse table from HTML to type Table
	parseHTMLTable(data, tbl)
	tbl.FindRange("test_table", true)
	tbl.Convert(true)

	ctx := context.Background()
	b, err := os.ReadFile("client_secret_445753725674-h9gtjqqf5fjfo3eljnr3gc893f77i4el.apps.googleusercontent.com.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	spreadsheetId := "1p43lzwwjdeH49N80yXcU7kfkg7RHdMin1K6Cqb0JH5Q"
	googleShtRange := tbl.sheetRange

	rb := &sheets.ValueRange{
		MajorDimension: "COLUMNS",
		Range:          googleShtRange,
		Values:         tbl.converted,
	}

	_, err = srv.Spreadsheets.Values.Update(spreadsheetId, googleShtRange, rb).ValueInputOption("RAW").Context(ctx).Do()
	if err != nil {
		log.Fatal(err)
	}
	//Test for make google sheet
	//rb := &sheets.Spreadsheet{
	//	Properties: &sheets.SpreadsheetProperties{
	//		Title: "HF_Test",
	//	},
	//}
	//_, err = srv.Spreadsheets.Create(rb).Context(ctx).Do()
	//if err != nil {
	//	log.Fatal(err)
	//}
}
