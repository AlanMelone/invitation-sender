package main

import (
	"flag"
	"fmt"

	"github.com/alanmelone/invitation-sender/internal/app/invitation_sender"
)

func main() {
	tablePathPtr := flag.String("t", "", "table of values")
	sheetNamePtr := flag.String("s", "", "sheet name")
	templatePath := flag.String("dt", "", "template for a message")

	flag.Parse()

	if *tablePathPtr == "" {
		panic("You should use -t flag for point to table")
	}
	fmt.Println("Table Path:", *tablePathPtr)

	_, err := invitation_sender.SendEmailFromTemplate(*tablePathPtr, *sheetNamePtr, *templatePath)
	if err != nil {
		panic(err)
	}
}
