import marimo

__generated_with = "0.17.4"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo
    return (mo,)


@app.cell
def _(mo):
    mo.md(
        """
        # PowerShell Learning

        This notebook provides interactive lessons for learning PowerShell.

        PowerShell is a task automation framework consisting of a command-line shell and
        scripting language built on .NET. It's used for system administration and automation.

        **Note**: This notebook demonstrates PowerShell commands and concepts.
        You can run PowerShell commands from Python using subprocess.
        """
    )
    return


@app.cell
def _():
    import subprocess
    import pandas as pd

    def run_powershell(command):
        """Execute a PowerShell command and return the output"""
        try:
            result = subprocess.run(
                ['pwsh', '-Command', command],
                capture_output=True,
                text=True,
                timeout=10
            )
            return result.stdout if result.returncode == 0 else f"Error: {result.stderr}"
        except subprocess.TimeoutExpired:
            return "Error: Command timed out"
        except FileNotFoundError:
            return "PowerShell (pwsh) not found. Install PowerShell Core or use 'powershell' on Windows."
        except Exception as e:
            return f"Error: {str(e)}"

    return run_powershell, subprocess, pd


@app.cell
def _(mo):
    mo.md("""
    ## 1. PowerShell Basics

    PowerShell commands are called **cmdlets** (pronounced "command-lets").
    They follow a `Verb-Noun` naming convention.

    **Common Verbs:**
    - `Get` - Retrieve information
    - `Set` - Change configuration
    - `New` - Create something
    - `Remove` - Delete something
    - `Start` - Begin a process
    - `Stop` - End a process
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Getting Help

    PowerShell has excellent built-in help:

    ```powershell
    Get-Help Get-Process
    Get-Help Get-Process -Examples
    Get-Command *Process*
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Example PowerShell Commands

    Here are some fundamental PowerShell commands you can try:
    """)
    return


@app.cell
def _(mo):
    # Example commands as code blocks
    examples = """
    # Get PowerShell version
    $PSVersionTable

    # List files in current directory
    Get-ChildItem

    # Get running processes
    Get-Process

    # Get services
    Get-Service

    # Get current location (directory)
    Get-Location

    # Get environment variables
    Get-ChildItem Env:
    """

    mo.md(f"""
    **Common PowerShell Commands:**

    ```powershell
    {examples}
    ```
    """)
    return (examples,)


@app.cell
def _(mo):
    mo.md("""
    ## 2. Variables and Data Types

    PowerShell variables start with `$` and are dynamically typed.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Variable Examples

    ```powershell
    # String
    $name = "PowerShell"

    # Number
    $count = 42

    # Boolean
    $isActive = $true

    # Array
    $colors = @("Red", "Green", "Blue")

    # Hash Table (Dictionary)
    $person = @{
        Name = "John"
        Age = 30
        City = "Seattle"
    }

    # Get variable type
    $name.GetType()
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### String Operations

    ```powershell
    # String concatenation
    $firstName = "John"
    $lastName = "Doe"
    $fullName = "$firstName $lastName"

    # String methods
    $text = "PowerShell"
    $text.Length
    $text.ToUpper()
    $text.ToLower()
    $text.Substring(0, 5)
    $text.Replace("Shell", "Tool")

    # String comparison
    "PowerShell" -eq "powershell"  # Case-insensitive by default
    "PowerShell" -ceq "powershell" # Case-sensitive
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 3. Working with Objects

    Everything in PowerShell is an object. Objects have properties and methods.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Object Examples

    ```powershell
    # Get object properties
    Get-Process | Get-Member

    # Select specific properties
    Get-Process | Select-Object Name, CPU, Memory

    # Create custom object
    $computer = [PSCustomObject]@{
        Name = "SERVER01"
        OS = "Windows Server 2022"
        RAM = 16
        Disks = 2
    }

    # Access properties
    $computer.Name
    $computer.RAM
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 4. Pipeline

    The pipeline (`|`) passes output from one command to another.
    This is one of PowerShell's most powerful features.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Pipeline Examples

    ```powershell
    # Get processes, sort by CPU usage, select top 5
    Get-Process | Sort-Object CPU -Descending | Select-Object -First 5

    # Get services that are running
    Get-Service | Where-Object Status -eq "Running"

    # Get files larger than 1MB
    Get-ChildItem | Where-Object Length -gt 1MB

    # Count files in directory
    Get-ChildItem | Measure-Object

    # Get total size of files
    Get-ChildItem | Measure-Object -Property Length -Sum
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 5. Filtering and Selecting

    Use `Where-Object` to filter and `Select-Object` to choose properties.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Filtering Examples

    ```powershell
    # Where-Object (full syntax)
    Get-Process | Where-Object { $_.CPU -gt 10 }

    # Where-Object (simplified)
    Get-Process | Where-Object CPU -gt 10

    # Multiple conditions
    Get-Service | Where-Object { $_.Status -eq "Running" -and $_.StartType -eq "Automatic" }

    # Select specific properties
    Get-Process | Select-Object Name, Id, CPU

    # Select and rename properties
    Get-Process | Select-Object @{Name='ProcessName';Expression={$_.Name}}, CPU

    # Select first/last items
    Get-Process | Select-Object -First 10
    Get-Process | Select-Object -Last 5
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 6. Arrays and Collections

    PowerShell provides powerful array and collection manipulation.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Array Examples

    ```powershell
    # Create array
    $numbers = 1, 2, 3, 4, 5
    $colors = @("Red", "Green", "Blue")

    # Array operations
    $numbers.Length
    $numbers[0]           # First element
    $numbers[-1]          # Last element
    $numbers[1..3]        # Range (elements 1-3)

    # Add to array
    $colors += "Yellow"

    # Array methods
    $numbers -contains 3
    $numbers -join ", "

    # ForEach loop
    foreach ($num in $numbers) {
        Write-Host $num
    }

    # ForEach-Object in pipeline
    $numbers | ForEach-Object { $_ * 2 }
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 7. Hash Tables

    Hash tables store key-value pairs (like dictionaries in Python).
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Hash Table Examples

    ```powershell
    # Create hash table
    $person = @{
        FirstName = "John"
        LastName = "Doe"
        Age = 30
        City = "Seattle"
    }

    # Access values
    $person["FirstName"]
    $person.FirstName

    # Add new key-value
    $person["Email"] = "john@example.com"
    $person.Phone = "555-0100"

    # Get all keys
    $person.Keys

    # Get all values
    $person.Values

    # Check if key exists
    $person.ContainsKey("Age")

    # Remove key
    $person.Remove("Phone")
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 8. Conditional Statements

    PowerShell supports standard conditional logic.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### If/Else Examples

    ```powershell
    # If statement
    $number = 10
    if ($number -gt 5) {
        Write-Host "Number is greater than 5"
    }

    # If-Else
    if ($number -eq 10) {
        Write-Host "Number is 10"
    } else {
        Write-Host "Number is not 10"
    }

    # If-ElseIf-Else
    $score = 85
    if ($score -ge 90) {
        Write-Host "Grade: A"
    } elseif ($score -ge 80) {
        Write-Host "Grade: B"
    } elseif ($score -ge 70) {
        Write-Host "Grade: C"
    } else {
        Write-Host "Grade: F"
    }

    # Switch statement
    $day = "Monday"
    switch ($day) {
        "Monday" { "Start of work week" }
        "Friday" { "End of work week" }
        "Saturday" { "Weekend!" }
        "Sunday" { "Weekend!" }
        default { "Midweek day" }
    }
    ```

    **Comparison Operators:**
    - `-eq` Equal to
    - `-ne` Not equal to
    - `-gt` Greater than
    - `-ge` Greater than or equal to
    - `-lt` Less than
    - `-le` Less than or equal to
    - `-like` Wildcard comparison
    - `-match` Regex match
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 9. Loops

    PowerShell provides several loop constructs.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Loop Examples

    ```powershell
    # For loop
    for ($i = 0; $i -lt 5; $i++) {
        Write-Host "Iteration $i"
    }

    # ForEach loop
    $colors = @("Red", "Green", "Blue")
    foreach ($color in $colors) {
        Write-Host "Color: $color"
    }

    # While loop
    $count = 0
    while ($count -lt 5) {
        Write-Host "Count: $count"
        $count++
    }

    # Do-While loop
    $num = 0
    do {
        Write-Host "Number: $num"
        $num++
    } while ($num -lt 3)

    # ForEach-Object (in pipeline)
    1..5 | ForEach-Object { $_ * 2 }

    # Break and Continue
    foreach ($i in 1..10) {
        if ($i -eq 5) { continue }  # Skip 5
        if ($i -eq 8) { break }     # Stop at 8
        Write-Host $i
    }
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 10. Functions

    Functions encapsulate reusable code blocks.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Function Examples

    ```powershell
    # Simple function
    function Get-Greeting {
        Write-Host "Hello, PowerShell!"
    }

    # Function with parameters
    function Get-PersonalGreeting {
        param(
            [string]$Name
        )
        Write-Host "Hello, $Name!"
    }

    # Function with multiple parameters and return value
    function Add-Numbers {
        param(
            [int]$Number1,
            [int]$Number2
        )
        return $Number1 + $Number2
    }

    # Advanced function with parameter validation
    function Get-ComputerInfo {
        param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$ComputerName,

            [ValidateSet("OS", "CPU", "Memory", "All")]
            [string]$InfoType = "All"
        )

        Write-Host "Getting $InfoType info for $ComputerName"
    }

    # Call functions
    Get-Greeting
    Get-PersonalGreeting -Name "Alice"
    $sum = Add-Numbers -Number1 5 -Number2 10
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 11. Working with Files and Directories

    PowerShell provides powerful file system manipulation.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### File Operations

    ```powershell
    # List files and directories
    Get-ChildItem
    Get-ChildItem -Path C:\\Users -Recurse

    # Filter by extension
    Get-ChildItem -Filter *.txt
    Get-ChildItem -Path . -Include *.ps1, *.psm1 -Recurse

    # Create directory
    New-Item -Path "C:\\Temp\\NewFolder" -ItemType Directory

    # Create file
    New-Item -Path "C:\\Temp\\test.txt" -ItemType File

    # Read file content
    Get-Content -Path "file.txt"

    # Write to file
    "Hello World" | Out-File -FilePath "output.txt"
    "Another line" | Add-Content -FilePath "output.txt"

    # Copy file
    Copy-Item -Path "source.txt" -Destination "destination.txt"

    # Move file
    Move-Item -Path "file.txt" -Destination "C:\\Temp\\file.txt"

    # Delete file
    Remove-Item -Path "file.txt"

    # Test if path exists
    Test-Path -Path "C:\\Temp\\file.txt"

    # Get file properties
    Get-Item -Path "file.txt" | Select-Object Name, Length, LastWriteTime
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 12. Working with CSV and JSON

    PowerShell makes it easy to work with structured data.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### CSV Examples

    ```powershell
    # Export to CSV
    Get-Process | Select-Object Name, CPU, Memory | Export-Csv -Path "processes.csv" -NoTypeInformation

    # Import from CSV
    $data = Import-Csv -Path "data.csv"

    # Work with CSV data
    $data | Where-Object Age -gt 30
    $data | Select-Object Name, Email
    ```

    ### JSON Examples

    ```powershell
    # Create object and convert to JSON
    $person = @{
        Name = "John Doe"
        Age = 30
        City = "Seattle"
    }
    $json = $person | ConvertTo-Json

    # Save JSON to file
    $person | ConvertTo-Json | Out-File "person.json"

    # Read and parse JSON
    $jsonContent = Get-Content "person.json" -Raw
    $data = $jsonContent | ConvertFrom-Json
    $data.Name
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 13. Error Handling

    PowerShell provides try-catch-finally for error handling.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Error Handling Examples

    ```powershell
    # Try-Catch
    try {
        Get-Content "nonexistent.txt" -ErrorAction Stop
    } catch {
        Write-Host "Error: $_"
    }

    # Try-Catch-Finally
    try {
        # Some risky operation
        $file = Get-Content "file.txt"
    } catch [System.IO.FileNotFoundException] {
        Write-Host "File not found"
    } catch {
        Write-Host "An error occurred: $_"
    } finally {
        Write-Host "Cleanup operations"
    }

    # Error preference
    $ErrorActionPreference = "Stop"  # Stop on errors
    $ErrorActionPreference = "Continue"  # Continue on errors (default)
    $ErrorActionPreference = "SilentlyContinue"  # Suppress errors
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 14. Remote Management

    PowerShell excels at managing remote computers.
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### Remote Management Examples

    ```powershell
    # Run command on remote computer
    Invoke-Command -ComputerName Server01 -ScriptBlock {
        Get-Process
    }

    # Run command on multiple computers
    Invoke-Command -ComputerName Server01, Server02, Server03 -ScriptBlock {
        Get-Service
    }

    # Start interactive session
    Enter-PSSession -ComputerName Server01

    # Exit interactive session
    Exit-PSSession

    # Create persistent session
    $session = New-PSSession -ComputerName Server01
    Invoke-Command -Session $session -ScriptBlock { Get-Process }
    Remove-PSSession $session
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ## 15. Useful Cmdlets Reference

    Here's a quick reference of commonly used cmdlets:
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ### File System
    - `Get-ChildItem` (alias: `ls`, `dir`) - List files/directories
    - `Set-Location` (alias: `cd`) - Change directory
    - `Get-Location` (alias: `pwd`) - Current directory
    - `New-Item` - Create file/directory
    - `Copy-Item` (alias: `cp`) - Copy files
    - `Move-Item` (alias: `mv`) - Move files
    - `Remove-Item` (alias: `rm`) - Delete files

    ### Process Management
    - `Get-Process` (alias: `ps`) - List processes
    - `Start-Process` - Start a process
    - `Stop-Process` (alias: `kill`) - Stop a process

    ### Service Management
    - `Get-Service` - List services
    - `Start-Service` - Start a service
    - `Stop-Service` - Stop a service
    - `Restart-Service` - Restart a service

    ### Data Manipulation
    - `Select-Object` - Select properties
    - `Where-Object` (alias: `where`, `?`) - Filter objects
    - `ForEach-Object` (alias: `foreach`, `%`) - Process each object
    - `Sort-Object` (alias: `sort`) - Sort objects
    - `Group-Object` - Group objects
    - `Measure-Object` - Calculate statistics

    ### Output
    - `Write-Host` - Display to console
    - `Write-Output` (alias: `echo`) - Send to pipeline
    - `Out-File` - Write to file
    - `Export-Csv` - Export to CSV
    - `ConvertTo-Json` - Convert to JSON
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ---

    ## Practice Area

    Use the interactive PowerShell executor below to try commands!

    **Note**: This requires PowerShell Core (pwsh) to be installed on your system.
    - macOS/Linux: `brew install powershell` or download from Microsoft
    - Windows: Usually pre-installed, or install PowerShell Core
    """)
    return


@app.cell
def _(mo):
    # Interactive PowerShell command executor
    ps_command = mo.ui.text_area(
        label="Enter PowerShell command:",
        placeholder="Get-Process | Select-Object -First 5",
        value="$PSVersionTable.PSVersion"
    )
    ps_command
    return (ps_command,)


@app.cell
def _(mo, run_powershell, ps_command):
    # Execute PowerShell command
    if ps_command.value:
        output = run_powershell(ps_command.value)
        mo.md(f"""
        **Command:**
        ```powershell
        {ps_command.value}
        ```

        **Output:**
        ```
        {output}
        ```
        """)
    else:
        mo.md("Enter a PowerShell command above to see the output")
    return (output,)


@app.cell
def _(mo):
    mo.md("""
    ---

    ## Quick Reference Card

    ### Variable Assignment
    ```powershell
    $name = "value"
    ```

    ### Comparison Operators
    ```
    -eq  (equal)
    -ne  (not equal)
    -gt  (greater than)
    -lt  (less than)
    -ge  (greater or equal)
    -le  (less or equal)
    -like (wildcard match)
    -match (regex match)
    ```

    ### Logical Operators
    ```
    -and
    -or
    -not (or !)
    ```

    ### Pipeline
    ```powershell
    Command1 | Command2 | Command3
    ```

    ### Common Patterns
    ```powershell
    # Get help
    Get-Help <cmdlet-name>

    # Find commands
    Get-Command *keyword*

    # Filter
    Get-Process | Where-Object CPU -gt 10

    # Select properties
    Get-Service | Select-Object Name, Status

    # Sort
    Get-ChildItem | Sort-Object Length -Descending

    # Export
    Get-Process | Export-Csv processes.csv
    ```
    """)
    return


@app.cell
def _(mo):
    mo.md("""
    ---

    ## Learning Resources

    - **Microsoft Learn**: Official PowerShell documentation
    - **PowerShell Gallery**: Community modules and scripts
    - **GitHub**: Many PowerShell examples and projects
    - **SS64.com**: PowerShell command reference

    ## Tips

    1. Use `Get-Help` liberally - PowerShell has excellent built-in documentation
    2. Tab completion is your friend - use it to discover cmdlets and parameters
    3. Use aliases for speed, but write scripts with full cmdlet names for clarity
    4. The pipeline is powerful - learn to chain commands together
    5. Objects, not text - PowerShell works with objects, making it more powerful than traditional shells
    """)
    return


if __name__ == "__main__":
    app.run()
