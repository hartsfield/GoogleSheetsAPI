package ohsheet

import (
	"context"
	"log"
	"os"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// Access contains values used for accessing the sheets API server, including
// access token, API credentials, and scopes
type Access struct {
	// Token is the path to your token.json file. If you're having trouble
	// authenticating, try deleting this file and running the program
	// again. This should renew your token. If you've never run this
	// program, you may not have a token. This program will generate a
	// token for you.
	Token string
	// Credentials is the path to your credentials.json file. This file can
	// be obtained from the API Keys section on Google Cloud Platform. You
	// may need to generate the file and enable the sheets API from within
	// Googl Cloud Platform.
	Credentials string
	// Scopes define what level(s) of access we'll have to the spreadsheet
	// If modifying these scopes, delete your previously saved token.json.
	Scopes []string
}

// Connect is used to connect to the sheets API
func (b *Access) Connect() *sheets.Service {
	ctx := context.Background()
	cred, err := os.ReadFile(b.Credentials)
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	config, err := google.ConfigFromJSON(cred, b.Scopes...)
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config, b.Token)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}
	return srv
}

func (b *Access) Write(srv *sheets.Service, spreadsheetId string, writeRange string, vals []interface{}) {
	var vr sheets.ValueRange
	vr.Values = append(vr.Values, vals)

	_, err := srv.Spreadsheets.Values.Update(spreadsheetId, writeRange, &vr).ValueInputOption("RAW").Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet. %v", err)
	}
}

func (b *Access) Read(srv *sheets.Service, spreadsheetId string, readRange string) (*sheets.ValueRange, error) {
	return srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
}
