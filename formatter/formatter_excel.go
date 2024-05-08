package formatter

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// ExcelFormatter is struct defined for Excel Output use-case
type ExcelFormatter struct {
    config *Config
}

// Format the data to Excel and output it to an Excel file
func (f *ExcelFormatter) Format(td *TemplateData, templateContent string) (err error) {
    file := excelize.NewFile()
    sheetName := "Sheet1"

    // Create a style for center alignment with text wrapping
    styleWithWrap, err := file.NewStyle(&excelize.Style{
        Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
    })

    // Create a style for center alignment without text wrapping
    styleWithoutWrap, err := file.NewStyle(&excelize.Style{
        Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
    })

    // Set the column headers
    file.SetCellValue(sheetName, "A1", "Activo")
    file.SetCellValue(sheetName, "B1", "IP")
    file.SetCellValue(sheetName, "C1", "Puerto/Servicio")
    file.SetCellValue(sheetName, "D1", "Banner")
    file.SetCellStyle(sheetName, "A1", "D1", styleWithWrap)

    // Set a default width for the columns
    file.SetColWidth(sheetName, "A", "A", 30)
    file.SetColWidth(sheetName, "B", "B", 20)
    file.SetColWidth(sheetName, "C", "C", 20)
    file.SetColWidth(sheetName, "D", "D", 100)

    row := 2 // Start from row 2 for data

    // Group hosts by IP
    ipToHosts := make(map[string][]*Host)
    for i := range td.NMAPRun.Host {
        var host *Host = &td.NMAPRun.Host[i]
        ipAddress := host.JoinedAddresses("/")
        ipToHosts[ipAddress] = append(ipToHosts[ipAddress], host)
    }

    // Iterate over the grouped hosts
    for ipAddress, hosts := range ipToHosts {
        domainStartRow := row // Keep track of the start row for this domain
        ipStartRow := row // Keep track of the start row for this IP/Host
        for _, host := range hosts {
            // Skipping hosts that are down
            if td.OutputOptions.ExcelOptions.SkipDownHosts && host.Status.State != "up" {
                continue
            }

            // Set the Host value
            file.SetCellValue(sheetName, fmt.Sprintf("A%d", row), host.JoinedHostNames("/"))
            file.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("A%d", row), styleWithWrap)

            // Set the IP value
            file.SetCellValue(sheetName, fmt.Sprintf("B%d", row), ipAddress)
            file.SetCellStyle(sheetName, fmt.Sprintf("B%d", row), fmt.Sprintf("B%d", row), styleWithoutWrap)

            for j := range host.Port {
                var port *Port = &host.Port[j]

                // Create a string representation of the port
                portStr := fmt.Sprintf("%d/%s %s", port.PortID, port.Protocol, port.Service.Name)

                // Set the Service value
                file.SetCellValue(sheetName, fmt.Sprintf("C%d", row), portStr)

                // Merge Product, Version, and Extra Info into one cell called Banner
                banner := fmt.Sprintf("%s %s %s", port.Service.Product, port.Service.Version, port.Service.ExtraInfo)
                file.SetCellValue(sheetName, fmt.Sprintf("D%d", row), banner)
                file.SetCellStyle(sheetName, fmt.Sprintf("C%d", row), fmt.Sprintf("D%d", row), styleWithWrap)

                row++ // Increment row for next port
            }
            // Merge the cells for the IP/Host
            file.MergeCell(sheetName, fmt.Sprintf("B%d", ipStartRow), fmt.Sprintf("B%d", row-1))
        }
        // Merge the cells for the domain
        file.MergeCell(sheetName, fmt.Sprintf("A%d", domainStartRow), fmt.Sprintf("A%d", row-1))
    }

    // Save the Excel file
    err = file.SaveAs("nmap-output.xlsx")
    return err
}

func (f *ExcelFormatter) defaultTemplateContent() string {
    return HTMLSimpleTemplate
}