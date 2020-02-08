# requires -Version 5.0 
class Person
{ 
  [string]$FirstName
  [string]$LastName
  [int][ValidateRange(0,100)]$Age
  [DateTime]$Birthday
 
  # constructor
  Person([string]$FirstName, [string]$LastName, [DateTime]$Birthday)
  {
    # set object properties
    $this.FirstName = $FirstName
    $this.LastName = $LastName
    $this.Birthday = $Birthday
    # calculate person age
    $ticks = ((Get-Date) - $Birthday).Ticks
    $this.Age = (New-Object DateTime -ArgumentList $ticks).Year-1
  }
} 
$Person1 = [Person]::new('Tobias','Weltner','2000-02-03')
$Person2 = [Person]::new('Frank','Peterson','1976-04-12')
$Person3 = [Person]::new('Helen','Stewards','1987-11-19') 

$Person1
$Person2
$Person3