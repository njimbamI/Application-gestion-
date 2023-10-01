package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/fatih/color"
	"github.com/tealeg/xlsx"
)

// Structure pour représenter une tâche
type Task struct {
	Title       string
	Description string
	CreatedAt   time.Time
	IsDone      bool
}

// Liste de tâches
var taskList []Task

func main() {
	fmt.Println("Gestionnaire de tâches\n")

	// Ajoutez des tâches de démonstration (facultatif)
	taskList = append(taskList, Task{
		Title:       "Tâche 1",
		Description: "Description de la tâche 1",
		CreatedAt:   time.Now(),
		IsDone:      false,
	})

	taskList = append(taskList, Task{
		Title:       "Tâche 2",
		Description: "Description de la tâche 2",
		CreatedAt:   time.Now(),
		IsDone:      true,
	})

	for {
		displayTasks()

		fmt.Println("\nMenu :")
		fmt.Println("1. Ajouter une tâche")
		fmt.Println("2. Supprimer une tâche")
		fmt.Println("3. Marquer une tâche comme terminée")
		fmt.Println("4. Exporter en xlsx")
		fmt.Println("5. Quitter")
		fmt.Print("Choisissez une option : ")

		var choice int
		fmt.Scanln(&choice)

		switch choice {
		case 1:
			addTask()
		case 2:
			deleteTask()
		case 3:
			markTaskAsDone()
		case 4:
			exportToXLSX()
		case 5:
			fmt.Println("A BIENTOT !")
			os.Exit(0)
		default:
			fmt.Println("Option invalide.")
		}
	}
}

// Fonction pour afficher la liste des tâches
func displayTasks() {
	if len(taskList) == 0 {
		fmt.Println("Aucune tâche.")
		return
	}

	fmt.Println("Liste des tâches :")
	fmt.Printf("%-4s | %-20s | %-40s | %-20s | %-10s\n", "No.", "Titre", "Description", "Date de création", "Statut")
	fmt.Println("---------------------------------------------------------------------------------------------------")

	for i, task := range taskList {
		status := "À faire"
		statusColor := color.New(color.FgRed)
		if task.IsDone {
			status = "Terminée"
			statusColor = color.New(color.FgGreen)
		}

		titleColor := color.New(color.FgWhite)
		descriptionColor := color.New(color.FgWhite)
		createdAtColor := color.New(color.FgWhite)
		statusColorTitle := titleColor
		statusColorDescription := descriptionColor
		statusColorCreatedAt := createdAtColor
		statusColorNumber := color.New(color.FgWhite)

		if task.IsDone {
			statusColorTitle = color.New(color.FgGreen)
			statusColorDescription = color.New(color.FgGreen)
			statusColorCreatedAt = color.New(color.FgGreen)
			statusColorNumber = color.New(color.FgGreen)
		} else {
			statusColorTitle = color.New(color.FgRed)
			statusColorDescription = color.New(color.FgRed)
			statusColorCreatedAt = color.New(color.FgRed)
			statusColorNumber = color.New(color.FgRed)
		}

		fmt.Printf("%-4s | ", statusColorNumber.Sprintf("%-4d", i+1))
		statusColorTitle.Printf("%-20s", task.Title)
		fmt.Printf(" | ")
		statusColorDescription.Printf("%-40s", task.Description)
		fmt.Printf(" | ")
		statusColorCreatedAt.Printf("%-20s", task.CreatedAt.Format("2006-01-02 15:04:05"))
		fmt.Printf(" | ")
		statusColor.Println(status)
	}
}

// Fonction pour ajouter une tâche
func addTask() {
	var title, description string

	// Demander à l'utilisateur le titre et la description de la tâche
	fmt.Print("Titre de la tâche : ")
	fmt.Scanln(&title)
	fmt.Print("Description de la tâche : ")
	fmt.Scanln(&description)

	// Créer une nouvelle tâche
	task := Task{
		Title:       title,
		Description: description,
		CreatedAt:   time.Now(),
		IsDone:      false,
	}
	// Ajouter la tâche à la liste des tâches
	taskList = append(taskList, task)
	// Afficher un message de confirmation
	fmt.Println("Tâche ajoutée avec succès !")
}

// Fonction pour supprimer une tâche
func deleteTask() {
	fmt.Println("\nListe des tâches à supprimer :")
	for i, task := range taskList {
		fmt.Printf("%d. %s\n", i+1, task.Title)
	}

	var choice int
	fmt.Print("Choisissez le numéro de la tâche à supprimer : ")
	fmt.Scanln(&choice)

	if choice >= 1 && choice <= len(taskList) {
		taskList = append(taskList[:choice-1], taskList[choice:]...)
		fmt.Println("Tâche supprimée avec succès !")
	} else {
		fmt.Println("Numéro de tâche invalide.")
	}
}

// Fonction pour marquer une tâche comme terminée
func markTaskAsDone() {
	fmt.Println("\nListe des tâches à marquer comme terminées :")
	for i, task := range taskList {
		if !task.IsDone {
			fmt.Printf("%d. %s\n", i+1, task.Title)
		}
	}

	var choice int
	fmt.Print("Choisissez le numéro de la tâche à marquer comme terminée : ")
	fmt.Scanln(&choice)

	if choice >= 1 && choice <= len(taskList) {
		taskList[choice-1].IsDone = true
		fmt.Println("Tâche marquée comme terminée !")
	} else {
		fmt.Println("Numéro de tâche invalide.")
	}
}

// Fonction pour exporter les tâches en xlsx
func exportToXLSX() {
	if len(taskList) == 0 {
		fmt.Println("Aucune tâche à exporter.")
		return
	}

	fmt.Print("Entrez le nom du fichier xlsx : ")
	var fileName string
	fmt.Scanln(&fileName)

	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Tâches")
	if err != nil {
		log.Fatal(err)
	}

	// Style pour l'en-tête en fond vert
	headerStyle := xlsx.NewStyle()
	headerStyle.Fill = *xlsx.NewFill("solid", "00FF00", "00FF00")

	// En-têtes
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "Titre"
	cell.SetStyle(headerStyle) // Appliquer le style à la cellule

	cell = row.AddCell()
	cell.Value = "Description"
	cell.SetStyle(headerStyle) // Appliquer le style à la cellule

	cell = row.AddCell()
	cell.Value = "Date de création"
	cell.SetStyle(headerStyle) // Appliquer le style à la cellule

	cell = row.AddCell()
	cell.Value = "Statut"
	cell.SetStyle(headerStyle) // Appliquer le style à la cellule

	// Ajouter les tâches
	for _, task := range taskList {
		row := sheet.AddRow()
		row.AddCell().Value = task.Title
		row.AddCell().Value = task.Description
		row.AddCell().Value = task.CreatedAt.Format("2006-01-02 15:04:05")
		if task.IsDone {
			row.AddCell().Value = "Terminée"
		} else {
			row.AddCell().Value = "À faire"
		}
	}

	err = file.Save("data/" + fileName + ".xlsx")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Tâches exportées dans le fichier '%s.xlsx' avec succès !\n", fileName)
}
