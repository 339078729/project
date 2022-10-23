package main

import (
	"fmt"
	"liuqi/readmodel"
	"os"
)

func main() {
	//fmt.Println("Hello, World!")
	readmodel.ReadExcel("./xlsx/demo.xlsx")
	fmt.Printf("Press any key to exit...")
	b := make([]byte, 1)
	os.Stdin.Read(b)
}
