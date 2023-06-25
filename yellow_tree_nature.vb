Public Module RecipeModule
Sub Main()
    'Declare objects
    Dim mainHeading As New Label
    Dim subHeading As New Label
    Dim recipeName As New TextBox
    Dim categoryDropDown As New DropDownList
    Dim ingredientsList As New TextBox
    Dim instructions As New TextBox
    Dim submitBtn As New Button
    
    'Set mainHeading properties
    mainHeading.Text = "Welcome to the Cooking Blog!"
    mainHeading.AutoSize = True
    mainHeading.Font.Bold = True
    mainHeading.Font.Size = 20
    
    'Set subHeading properties
    subHeading.Text = "Please submit your recipe below:"
    subHeading.AutoSize = True
    subHeading.Font.Size = 13
    
    'Set recipeName properties
    recipeName.Name = "RecipeName"
    recipeName.Text = ""
    recipeName.Font.Size = 12
    recipeName.Width = 200
    
    'Set categoryDropDown properties
    categoryDropDown.Name = "Category"
    categoryDropDown.Items.Add("Breakfast")
    categoryDropDown.Items.Add("Lunch")
    categoryDropDown.Items.Add("Dinner")
    categoryDropDown.Items.Add("Snacks")
    categoryDropDown.Items.Add("Desserts")
    categoryDropDown.Width = 200
    categoryDropDown.Font.Size = 12
    categoryDropDown.SelectedIndex = 0
    
    'Set ingredientsList properties
    ingredientsList.Name = "Ingredients"
    ingredientsList.Text = ""
    ingredientsList.Font.Size = 12
    ingredientsList.Width = 200
    
    'Set instructions properties
    instructions.Name = "Instructions"
    instructions.Text = ""
    instructions.Font.Size = 12
    instructions.Width = 200
    
    'Set submitBtn properties
    submitBtn.Name = "Submit"
    submitBtn.Text = "Submit"
    submitBtn.Width = 100
    submitBtn.Font.Size = 12
    
    'Add controls to the form
    Form1.Controls.Add(mainHeading)
    Form1.Controls.Add(subHeading)
    Form1.Controls.Add(recipeName)
    Form1.Controls.Add(categoryDropDown)
    Form1.Controls.Add(ingredientsList)
    Form1.Controls.Add(instructions)
    Form1.Controls.Add(submitBtn)
    
    'Set positions for controls
    mainHeading.Location = New Point(50, 20)
    subHeading.Location = New Point(50, 50)
    recipeName.Location = New Point(50, 80)
    categoryDropDown.Location = New Point(50, 110)
    ingredientsList.Location = New Point(50, 140)
    instructions.Location = New Point(50, 170)
    submitBtn.Location = New Point(50, 200)
    
    'Set form properties
    Form1.Text = "Cooking and Recipe Blog"
    Form1.Height = 300
    Form1.Width = 350
    Form1.AutoSize = True
    Form1.ShowDialog()
End Sub

Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles Form1.Load
    'Show the main form
    Form1.Show()
End Sub

Private Sub submitBtn_Click(sender As Object, e As EventArgs) Handles submitBtn.Click
    'Declare variables 
    Dim recipe As String = recipeName.Text
    Dim category As String = categoryDropDown.SelectedItem.ToString()
    Dim ing As String = ingredientsList.Text
    Dim inst As String = instructions.Text
    
    'Declare and instantiate a new recipe object
    Dim newRecipe As Recipe = New Recipe(recipe, category, ing, inst)
    
    If (recipeName.Text = "" Or ingredientsList.Text = "" Or instructions.Text = "") Then
        MessageBox.Show("Please fill in all required fields")
    Else
        'Call the SaveRecipe method of the Recipe class
        newRecipe.SaveRecipe()
    End If
End Sub

End Module

Public Class Recipe

Public Sub New(ByVal Name As String, ByVal Category As String, ByVal Ingredients As String, ByVal Instructions As String)
    Me.Name = Name
    Me.Category = Category
    Me.Ingredients = Ingredients
    Me.Instructions = Instructions
End Sub

'Declare private variables 
Private Name As String
Private Category As String
Private Ingredients As String
Private Instructions As String

'Declare public properties
Public Property RecipeName As String
    Get
        Return Name
    End Get
    Set(ByVal value As String)
        Name = value
    End Set
End Property

Public Property RecipeCategory As String
    Get
        Return Category
    End Get
    Set(ByVal value As String)
        Category = value
    End Set
End Property

Public Property RecipeIngredients As String
    Get
        Return Ingredients
    End Get
    Set(ByVal value As String)
        Ingredients = value
    End Set
End Property

Public Property RecipeInstructions As String
    Get
        Return Instructions
    End Get
    Set(ByVal value As String)
        Instructions = value
    End Set
End Property

Public Sub SaveRecipe()
    'Format the recipe text
    Dim recipeText As String = String.Format("Recipe Name: {0}{1}Category: {2}{1}Ingredients: {3}{1}Instructions: {4}", Name, vbCrLf, Category, Ingredients, Instructions)
    
    'Write the recipe text to a file
    Dim filename As String = "C:\CookingBlog\Recipes.txt"
    Dim file As System.IO.StreamWriter
    file = My.Computer.FileSystem.OpenTextFileWriter(filename, True)
    file.WriteLine(recipeText)
    file.Close()
    
    'Display success message
    MessageBox.Show("Recipe Submitted Successfully!")
End Sub

End Class