package main

import (
	"context"
	"errors"
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"cloud.google.com/go/firestore"
	firebase "firebase.google.com/go"
	auth "firebase.google.com/go/auth"
	"github.com/tealeg/xlsx"
	"google.golang.org/api/option"
)

func readSheetToSliceOfMap(sheet *xlsx.Sheet) (res []map[string]string, err error) {
	headers := make([]string, 0)
	for i, row := range sheet.Rows {
		if i == 0 {
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					return nil, err
				}
				headers = append(headers, fmt.Sprintf("%s", str))
			}
			continue
		}
		vals := make(map[string]string)
		if row != nil {
			p := false
			for _, cell := range row.Cells {
				if cell.String() != "" {
					p = true
					break
				}
			}
			if !p {
				continue
			}
			for j, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					return nil, err
				}
				vals[headers[j]] = fmt.Sprintf("%s", str)
			}
			res = append(res, vals)
		}
	}

	return res, nil
}

func readFromSourceExcel(filename string) (userlines []map[string]string, projectlines []map[string]string, contactlines []map[string]string, err error) {
	var xlFile *xlsx.File

	xlFile, err = xlsx.OpenFile(filename)
	if err != nil {
		return nil, nil, nil, err
	}

	if len(xlFile.Sheets) == 0 {
		return nil, nil, nil, errors.New("This XLSX file contains no sheets")
	}

	userlines, err = readSheetToSliceOfMap(xlFile.Sheets[0])
	if err != nil {
		return nil, nil, nil, err
	}

	projectlines, err = readSheetToSliceOfMap(xlFile.Sheets[1])
	if err != nil {
		return nil, nil, nil, err
	}

	contactlines, err = readSheetToSliceOfMap(xlFile.Sheets[2])
	if err != nil {
		return nil, nil, nil, err
	}

	return userlines, projectlines, contactlines, nil
}

func main() {
	var xlsxPath string
	if len(os.Args) < 2 {
		xlsxPath = "upload_sheet.xlsx"
	} else {
		xlsxPath = os.Args[1]
	}
	log.Printf("Use %q file as data source \n", xlsxPath)

	userlines := make([]map[string]string, 0, 0)
	projectlines := make([]map[string]string, 0, 0)
	contactlines := make([]map[string]string, 0, 0)
	userlines, projectlines, contactlines, err := readFromSourceExcel(xlsxPath)
	if err != nil {
		log.Fatal(err)
	}

	opt := option.WithCredentialsFile("serviceAccountKey.json")
	ctx := context.Background()

	app, err := firebase.NewApp(ctx, nil, opt)
	if err != nil {
		log.Fatalf("Error initializing app: '%v'", err)
	}
	firestoreClient, err := app.Firestore(ctx)
	if err != nil {
		log.Fatalln(err)
	}
	defer firestoreClient.Close()

	if len(userlines) != 0 {
		fmt.Printf("Create user records:")
		authClient, err := app.Auth(ctx)
		if err != nil {
			log.Fatalf("Error getting Auth client: %v\n", err)
		}
		for _, line := range userlines {
			u, err := authClient.GetUserByEmail(ctx, line["identifier"])
			if err != nil {
				if !strings.Contains(err.Error(), "cannot find user from email") {
					log.Fatalf("Error getting user by email %s: %v\n", line["identifier"], err)
				}
			}
			if u != nil {
				continue
			}

			params := (&auth.UserToCreate{}).
				Email(line["identifier"]).
				EmailVerified(false).
				Password("12345678").
				Disabled(false)

			UserRecord, err := authClient.CreateUser(ctx, params)
			if err != nil {
				log.Fatalf("error creating user: %v\n", err)
			}

			_, err = firestoreClient.Collection("users").Doc(UserRecord.UID).Set(ctx, map[string]interface{}{
				"first_name": line["first_name"],
				"last_name":  line["last_name"],
				"role":       line["role"],
			}, firestore.MergeAll)

			fmt.Printf("Successfully created user: %q with uid: %q\n", line["identifier"], UserRecord.UID)
		}
	}

	var projectID string
	fmt.Println()
	fmt.Print("Add project details and cable log:")
	for _, line := range projectlines {
		fmt.Print(".")
		if line["project_id"] != "" {
			projectID = line["project_id"]
			date, _ := time.Parse("01/02/2006", line["start_date"])
			_, err = firestoreClient.Collection("project1").Doc(projectID).Set(ctx, map[string]interface{}{
				"address_line_1": line["address_line_1"],
				"address_line_2": line["address_line_2"],
				"area":           line["area"],
				"client_name":    line["client_name"],
				"contact_name":   line["contact_name"],
				"contact_phone":  line["contact_phone"],
				"engineer_id":    line["engineer_id"],
				"field_tech_id":  line["field_tech_id"],
				"map_image":      line["map_image"],
				"project_id":     projectID,
				"start_date":     date,
			}, firestore.MergeAll)
			if err != nil {
				log.Fatalf("Failed adding %v: %v", line, err)
			}
		}

		// TODO cable log
	}
	fmt.Println()
	fmt.Print("Add contacts:")
	for _, line := range contactlines {
		fmt.Print(".")
		if line["email"] != "" {
			_, _, err = firestoreClient.Collection("project1").Doc(projectID).Collection("contacts").Add(ctx, map[string]interface{}{
				"email":  line["email"],
				"name":   line["name"],
				"status": line["status"],
			})
			if err != nil {
				log.Fatalf("Failed adding %v: %v", line, err)
			}
		}
	}

	fmt.Println()
	log.Println("Job done!")
}
