// Package ohsheet is a Go library for reading and writing data to/from Google
// Sheets using the Google Sheets API

// Copyright (c) 2023 J. Hartsfield

// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:

// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
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

// Write is used to write to a spreadsheet
func (b *Access) Write(srv *sheets.Service, spreadsheetId string, writeRange string, vals []interface{}) (*sheets.UpdateValuesResponse, error) {
	var vr sheets.ValueRange
	vr.Values = append(vr.Values, vals)

	return srv.Spreadsheets.Values.Update(spreadsheetId, writeRange, &vr).ValueInputOption("RAW").Do()
}

// Read is used to read from a spreadsheet
func (b *Access) Read(srv *sheets.Service, spreadsheetId string, readRange string) (*sheets.ValueRange, error) {
	return srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
}
