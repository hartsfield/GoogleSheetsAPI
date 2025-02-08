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
	Context context.Context
	// Token is the path to your token.json file. If you're having trouble
	// authenticating, try deleting this file and running the program
	// again. This should renew your token. If you've never run this
	// program, you may not have a token. This program will generate a
	// token for you.
	Token string
	// Credentials is the path to your credentials.json file. This file can
	// be obtained from the API Keys section of Google Cloud Platform. You
	// may need to generate the file and enable the sheets API from within
	// Google Cloud Platform.
	Credentials string
	// Scopes define what level(s) of access we'll have to the spreadsheet
	// If modifying these scopes, delete your previously saved token.json.
	Scopes []string
}

// Connect is used to connect to the sheets API. The code for this was mostly
// written by Google devs (see: connect.go). It will check for your
// credentials.json file in the root directory.
//
// Obtaining credentials:
//
// You download the credentials.json file from Google Cloud console in the
// "APIs and Credentials" section. You may need to generate the file.
// This file is used along with your selected scopes to generate a config.
//
// Obtaining a token:
//
// NOTE: These instructions only apply when your credentials are generated
// for a stand-alone desktop app. Other types of apps may need to obtain a
// token by other means.
//
// If you don't have a token, this program will help you obtain one on the
// first run. If you start getting errors you may need to delete the token and
// restart the program to generate a new one. On the first run, start the
// program in a terminal, and you'll receive a link to open in your browser to
// allow the program access to your account. Once you allow access you'll
// arrive at a dead link that looks like this:
//
// http://localhost/?state=state-token&code=LONG_CODE_HERE_COPY_THIS&scope=https://www.googleapis.com/auth/spreadsheets
//
// Copy the code found between `code=` and `&scope`, and then come back to the
// terminal, paste it in, and press enter. This will download a fresh token,
// and your program will proceed as normal. You'll only need to do this on the
// first run and when your token expires.
func (a *Access) Connect() *sheets.Service {
	a.Context = context.Background()
	cred, err := os.ReadFile(a.Credentials)
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	config, err := google.ConfigFromJSON(cred, a.Scopes...)
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config, a.Token)
	srv, err := sheets.NewService(a.ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}
	return srv
}

// Write is used to write to a spreadsheet
func (a *Access) Write(srv *sheets.Service, spreadsheetId string, writeRange string, vals []interface{}) (*sheets.AppendValuesResponse, error) {
	var vr sheets.ValueRange
	vr.Values = append(vr.Values, vals)

	return srv.Spreadsheets.Values.Append(spreadsheetId, writeRange, &vr).ValueInputOption("RAW").Do()
}

// Read is used to read from a spreadsheet
func (a *Access) Read(srv *sheets.Service, spreadsheetId string, readRange string) (*sheets.ValueRange, error) {
	return srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
}
