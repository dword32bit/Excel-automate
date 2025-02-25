package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func fillPreventiveColumn(inputFile string) error {
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		return fmt.Errorf("failed to open Excel file: %w", err)
	}

	sheetName := f.GetSheetName(0)

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get rows from sheet: %w", err)
	}

	for i := 1; i < len(rows); i++ {

		label := rows[i][41]
		protocol := rows[i][1]
		service := rows[i][2]

		var preventive string

		if label == "normal" { //normal
			preventive = "allow"
		} else if label == "neptune" && protocol == "tcp" && service == "http" { //neptune
			preventive = "nep-tcp-http"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "neptune" && protocol == "tcp" && service == "private" {
			preventive = "nep-tcp-private"
		} else if label == "warezclient" {
			preventive = "deny"
		} else if label == "ipsweep" {
			preventive = "deny"
		} else if label == "portsweep" {
			preventive = "deny"
		} else if label == "teardrop" {
			preventive = "deny"
		} else if label == "nmap" {
			preventive = "deny"
		} else if label == "satan" {
			preventive = "deny"
		} else if label == "smurf" {
			preventive = "deny"
		} else if label == "pod" {
			preventive = "deny"
		} else if label == "back" {
			preventive = "deny"
		} else if label == "guess_passwd" {
			preventive = "deny"
		} else if label == "ftp_write" {
			preventive = "deny"
		} else if label == "multihop" {
			preventive = "deny"
		} else if label == "rootkit" {
			preventive = "deny"
		} else if label == "buffer_overflow" {
			preventive = "deny"
		} else if label == "imap" {
			preventive = "deny"
		} else if label == "warezmaster" {
			preventive = "deny"
		} else if label == "phf" {
			preventive = "deny"
		} else if label == "land" {
			preventive = "deny"
		} else if label == "loadmodule" {
			preventive = "deny"
		} else if label == "spy" {
			preventive = "deny"
		} else if label == "perl" {
			preventive = "deny"
		} else if label == "saint" {
			preventive = "deny"
		} else if label == "mscan" {
			preventive = "deny"
		} else if label == "apache2" {
			preventive = "deny"
		} else if label == "snmpgetattack" {
			preventive = "deny"
		} else if label == "processtable" {
			preventive = "deny"
		} else if label == "httptunnel" {
			preventive = "deny"
		} else if label == "ps" {
			preventive = "deny"
		} else if label == "snmpguess" {
			preventive = "deny"
		} else if label == "mailbomb" {
			preventive = "deny"
		} else if label == "named" {
			preventive = "deny"
		} else if label == "sendmail" {
			preventive = "deny"
		} else if label == "xterm" {
			preventive = "deny"
		} else if label == "worm" {
			preventive = "deny"
		} else if label == "xlock" {
			preventive = "deny"
		} else if label == "snoop" {
			preventive = "deny"
		} else if label == "sqlattack" {
			preventive = "deny"
		} else if label == "udpstorm" {
			preventive = "deny"
		} else {
			preventive = "" // Set to empty if label not found
		}

		// Write the preventive value to column 43 (index 42)
		rows[i] = append(rows[i][:42], preventive)

		// debug
		// fmt.Printf("Row %d: Protocol = %s, Service = %s\n", i+1, protocol, service)
	}

	for i, row := range rows {
		cell, err := excelize.CoordinatesToCellName(1, i+1)
		if err != nil {
			return fmt.Errorf("failed to convert coordinates to cell name: %w", err)
		}
		if err := f.SetSheetRow(sheetName, cell, &row); err != nil {
			return fmt.Errorf("failed to set sheet row: %w", err)
		}
	}

	if err := f.Save(); err != nil {
		return fmt.Errorf("failed to save Excel file: %w", err)
	}

	return nil
}

func main() {
	if err := fillPreventiveColumn("contoh.xlsx"); err != nil {
		log.Fatalf("Error: %v", err)
	}
	fmt.Println("Successfully updated the Excel file.")
}
