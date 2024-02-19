package formatter

import (
	"fmt"
	"strings"

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
    file.SetCellValue(sheetName, "A1", "Dominio")
    file.SetCellValue(sheetName, "B1", "IP/Host")
    file.SetCellValue(sheetName, "C1", "Servicios")
    file.SetCellStyle(sheetName, "A1", "A1", styleWithWrap)
    file.SetCellStyle(sheetName, "B1", "B1", styleWithoutWrap)
    file.SetCellStyle(sheetName, "C1", "C1", styleWithWrap)

    // Set a default width for the columns
    file.SetColWidth(sheetName, "A", "C", 20)

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
        // Set the IP value
        file.SetCellValue(sheetName, fmt.Sprintf("B%d", row), ipAddress)
        file.SetCellStyle(sheetName, fmt.Sprintf("B%d", row), fmt.Sprintf("B%d", row), styleWithoutWrap)

        // Combine all host names for this IP into a single string
        var hostNames []string
        for _, host := range hosts {
            hostNames = append(hostNames, host.JoinedHostNames("/"))
        }
        hostNamesStr := strings.Join(hostNames, "\n")

        // Set the Host value
        file.SetCellValue(sheetName, fmt.Sprintf("A%d", row), hostNamesStr)
        file.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("A%d", row), styleWithWrap)

        // Keep track of the ports that have been added for this IP
        addedPorts := make(map[string]bool)
        var portStrings []string

        for _, host := range hosts {
            // Skipping hosts that are down
            if td.OutputOptions.ExcelOptions.SkipDownHosts && host.Status.State != "up" {
                continue
            }

            for j := range host.Port {
                var port *Port = &host.Port[j]

                // Create a string representation of the port
                portStr := fmt.Sprintf("%d/%s %s", port.PortID, port.Protocol, port.Service.Name)

                // Skip this port if it has already been added for this IP
                if addedPorts[portStr] {
                    continue
                }

                // Add this port to the list of ports for this IP
                portStrings = append(portStrings, portStr)

                // Mark this port as added for this IP
                addedPorts[portStr] = true
            }
        }

        // Combine all port strings for this IP into a single string
        portsStr := strings.Join(portStrings, "\n")

        // Set the Service value
        file.SetCellValue(sheetName, fmt.Sprintf("C%d", row), portsStr)
        file.SetCellStyle(sheetName, fmt.Sprintf("C%d", row), fmt.Sprintf("C%d", row), styleWithWrap)

        row++ // Increment row for next IP

    }

    // Save the Excel file
    err = file.SaveAs("nmap-output.xlsx")
    return err
}

func (f *ExcelFormatter) defaultTemplateContent() string {
	return HTMLSimpleTemplate
}
