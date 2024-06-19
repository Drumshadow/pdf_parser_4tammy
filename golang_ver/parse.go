package main

import (
    "bytes"
    "context"
    "fmt"
	"os"
    "os/exec"
	"strings"
)

func main() {
    // See "man pdftotext" for more options.
	pdf := os.Args[1]
	pdf_to_text_exe := os.Args[2]
    args := []string{
        "-layout",              // Maintain (as best as possible) the original physical layout of the text.
        "-nopgbrk",             // Don't insert page breaks (form feed characters) between pages.
        pdf, 					// The input file.
        "-",                    // Send the output to stdout.
    }
    cmd := exec.CommandContext(context.Background(), pdf_to_text_exe, args...)

    var buf bytes.Buffer
    cmd.Stdout = &buf

    if err := cmd.Run(); err != nil {
        fmt.Println(err)
        return
    }

    split_arr := strings.Split(buf.String(), "\n")
	fmt.Println(split_arr)

	for i := 0; i < len(split_arr); i++{
		if strings.Contains(split_arr[i], "Code Total"){
			fmt.Println(split_arr[i])
		}
	}

}