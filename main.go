package main

import (
	"fmt"
	"os"
)

type DBCFile struct {
	Version  string
	Nodes    []string
	Messages []Message
}

type Message struct {
	ID      uint32
	Name    string
	Size    uint8
	Node    string
	Signals []Signal
}

type Signal struct {
	Name      string
	StartBit  uint8
	Length    uint8
	Endianess string // "Intel" (little) или "Motorola" (big)
	Signed    bool
	Scale     float64
	Offset    float64
	Min       float64
	Max       float64
	Unit      string
	Receivers []string
}

func main() {
	dbc := DBCFile{
		Version: "1.0",
		Nodes:   []string{"ECU1", "ECU2"},
		Messages: []Message{
			{
				ID:   0x123,
				Name: "EngineData",
				Size: 8,
				Node: "ECU1",
				Signals: []Signal{
					{
						Name:      "RPM",
						StartBit:  0,
						Length:    16,
						Endianess: "Intel",
						Signed:    false,
						Scale:     0.25,
						Offset:    0,
						Min:       0,
						Max:       16383.75,
						Unit:     "rpm",
						Receivers: []string{"ECU2"},
					},
				},
			},
		},
	}

	file, err := os.Create("example.dbc")
	if err != nil {
		panic(err)
	}
	defer file.Close()

	_, err = file.WriteString(GenerateDBC(dbc))
	if err != nil {
		panic(err)
	}

	fmt.Println("DBC file created successfully!")
}

func GenerateDBC(dbc DBCFile) string {
	var content string

	content += "VERSION \"" + dbc.Version + "\"\n\n"

	content += "NS_ :\n"
	for _, node := range dbc.Nodes {
		content += "BU_ " + node + "\n"
	}
	content += "\n"
	
	for _, msg := range dbc.Messages {
		content += fmt.Sprintf("BO_ %d %s: %d %s\n", msg.ID, msg.Name, msg.Size, msg.Node)
		for _, sig := range msg.Signals {
			content += fmt.Sprintf(" SG_ %s : %d|%d@%s%+f (%f,%f) [%f|%f] \"%s\" ", 
				sig.Name, sig.StartBit, sig.Length, 
				map[string]string{"Intel": "1", "Motorola": "0"}[sig.Endianess],
				sig.Scale, sig.Offset, sig.Scale, sig.Min, sig.Max, sig.Unit)
			for i, receiver := range sig.Receivers {
				if i > 0 {
					content += ","
				}
				content += receiver
			}
			content += "\n"
		}
		content += "\n"
	}

	return content
}