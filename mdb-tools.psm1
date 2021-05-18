#-----------------------------------------------------------------------------------------
# MDB TOOLS
# Author: Mukunda Johnson (me@mukunda.com)
# License: MIT
#
# Handy functions for inspecting MDB database files and differences.
#-----------------------------------------------------------------------------------------

#-----------------------------------------------------------------------------------------
# Outputs any difference in the table listing.
Function Get-Tables-Diff {
   Param( $db1, $db2 )

   $tables1 = $db1.TableDefs | %{ $_.Name.ToLower() }
   $tables2 = $db2.TableDefs | %{ $_.Name.ToLower() }

   $tables1 | %{
      if( $tables2 -notcontains $_ ) {
         Write-Output "(Schema) DB1 has additional table $_"
         #[AdditionalTable]::new( "DB1", $_ )
      }
   }

   $tables2 | %{
      if( $tables1 -notcontains $_ ) {
         Write-Output "(Schema) DB2 has additional table $_"
         #[AdditionalTable]::new( "DB2", $_ )
      }
   }
}

#-----------------------------------------------------------------------------------------
# Intersects the table list and returns a list of names that are present in both
#  tables.
function Select-Common-Tables {
   Param( $db1, $db2 )
   $tables1 = $db1.TableDefs | %{ $_.Name }
   $tables2 = $db2.TableDefs | %{ $_.Name }
   
   $tables1 | Where-Object {$tables2 -contains $_}
}

#-----------------------------------------------------------------------------------------
# Checks field list in both tables and outputs any differences.
function Get-Table-Fields-Diff {
   Param( $table, $db1, $db2 )
   $table1 = $db1.TableDefs | Where-Object {$_.Name -eq $table}
   $table2 = $db2.TableDefs | Where-Object {$_.Name -eq $table}

   $not_found = [System.Collections.ArrayList]@(($table2.Fields | %{ $_.Name.ToLower() }))
   
   foreach( $ca in $table1.Fields ) {
      $fieldname = $ca.Name.ToLower()
      $cb = $table2.Fields | Where-Object {$_.Name.ToLower() -eq $fieldname}
      if( -not $cb ) {
         Write-Output "DB1 has additional field in ${table}: $fieldname"
         continue
      }

      $not_found.Remove( $fieldname )

      # AllowZeroLength seems to not be handled properly by the ODBC connector.
      $comparisons = ("Type", "Size", "DefaultValue", "Required")#, "AllowZeroLength")
      $comparisons | %{
         if( $ca.$_ -ne $cb.$_ ) {
            Write-Output "$table.$fieldname - `"$_`" differs: $($ca.$_) - $($cb.$_)"
         }
      }
   }
   
   foreach( $a in $not_found ) {
      Write-Output "DB2 has additional field in ${table}: $a"
   }
}

#-----------------------------------------------------------------------------------------
# Returns a list of field names that are present in both tables.
function Select-Common-Fields {
   Param( $table, $db1, $db2 )
   $fields1 = @(($db1.TableDefs | Where-Object {$_.Name -eq $table}).Fields | %{$_.Name})
   $fields2 = @(($db2.TableDefs | Where-Object {$_.Name -eq $table}).Fields | %{$_.Name})

   $fields1 | Where-Object {$fields2 -contains $_}
}

#-----------------------------------------------------------------------------------------
class RowData {
   [int]$RowIndex
   [string[]]$Data
   RowData( [int]$RowIndex, [string[]]$Data ) {
      $this.RowIndex      = $RowIndex
      $this.Data          = $Data
   }
}

#-----------------------------------------------------------------------------------------
class RowDifferences {
   [string[]]$Name
   [string[]]$Fields
   [RowData[]]$Rows1
   [RowData[]]$Rows2
   [boolean]$Truncated
   RowDifferences( [string]$Name, [string[]]$Fields ) {
      $this.Name   = $Name
      $this.Fields = $Fields
   }

   #--------------------------------------------------------------------------------------
   # Using the 5.1 format xml file is so nasty, and so is making classes for every output
   #  type. Sure, that's more flexible, but nobody is really going to use this script for
   #  more than some simple preliminary checks before closer inspection with MS Access.
   #
   # So, we're just going to use this to print output as strings.
   # The other functions also just output strings.
   [string]Print() {
      
      $a = $this
      $out = ""

      $fieldSizes = @()
      for( $i = 0; $i -lt $a.Fields.Count; $i++ ) {
         $fieldSize = $a.Fields[$i].Length
         $maxSize = 15
         if( $maxSize -lt $fieldSize ) { $maxSize = $fieldSize }

         for( $j = 0; $j -lt $a.Rows1.Count; $j++ ) {
            if( $a.Rows1[$j].Data[$i].Length -gt $fieldSize ) {
               $fieldSize = $a.Rows1[$j].Data[$i].Length
            }
         }
         for( $j = 0; $j -lt $a.Rows2.Count; $j++ ) {
            if( $a.Rows2[$j].Data[$i].Length -gt $fieldSize ) {
               $fieldSize = $a.Rows2[$j].Data[$i].Length
            }
         }

         if( $fieldSize -gt $maxSize ) { $fieldSize = $maxSize }
         $fieldSizes += $fieldSize
      }  

      $out += "[Differences in $($a.Name)]`n"
      $out += "   Row"
      for( $i = 0; $i -lt $a.Fields.Count; $i++ ) {
         $formatted = $a.Fields[$i]
         if( $formatted.Length -gt $fieldSizes[$i] ) {
            $formatted = $formatted.SubString( 0, $fieldSizes[$i] )
         } else {
            $formatted = $formatted.PadLeft( $fieldSizes[$i], " " )
         }
         
         $out += (" " + $formatted)
      }
      $out += "`n"
      $out += "------"
      for( $i = 0; $i -lt $a.Fields.Count; $i++ ) {
         $out += (" " + "".PadLeft( $fieldSizes[$i], "-" ))
      }
        
      for( $i = 0; $i -lt $a.Rows1.Count; $i++ ) {
         $out += "`n"

         $out += (([string]$a.Rows1[$i].RowIndex).PadLeft(6))
         for( $j = 0; $j -lt $a.Fields.Count; $j++ ) {
            $diff = $a.Rows1[$i].Data[$j] -ne $a.Rows2[$i].Data[$j]
            
            $v = [string]$a.Rows1[$i].Data[$j]
            if( $v.Length -gt $fieldSizes[$j] ) { $v = $v.Substring( 0, $fieldSizes[$j] ) }
            $v = $v.PadLeft( $fieldSizes[$j], ' ' )
            if( $diff ) { $v = "$([char]27)[5;93m$v$([char]27)[0m" }
            $out += (" $v")
         }

         $out += "`n"
         $out += "      "
         for( $j = 0; $j -lt $a.Fields.Count; $j++ ) {
            $diff = $a.Rows2[$i].Data[$j] -ne $a.Rows1[$i].Data[$j]
            
            $v = [string]$a.Rows2[$i].Data[$j]
            if( $v.Length -gt $fieldSizes[$j] ) { $v = $v.Substring( 0, $fieldSizes[$j] ) }
            $v = $v.PadLeft( $fieldSizes[$j], ' ' )
            if( $diff ) { $v = "$([char]27)[5;93m$v$([char]27)[0m" }
            $out += (" $v")
         }
      }
        
      $out += "`n"
      if( $a.Truncated ) {
         $out += "---this result was truncated---`n"
      }
      
      return $out
   }
}

Function Get-MDBDiff {
   <#
      .SYNOPSIS
         Compares two MS Access database files (MDB) and prints differences.

      .EXAMPLE
         Get-MDBDiff .\database1.mdb .\database2.mdb

         Scans the two MDB files and prints a report of differences.
   #>
   Param(
      [string]$Path1,
      [string]$Path2
   )

   # Initialize DB Engine
   $dbe = New-Object -comobject DAO.DBEngine.120

   # Open databases. Note these need absolute paths.
   # Params are path, "options", and "read-only"
   # I want to try and get around any locks on the file, but I'm not sure if this gets
   #  past an exclusive lock.
   # "Options" controls exclusive mode.
   # https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-opendatabase-method-dao
   $db1 = $dbe.OpenDatabase( (Resolve-Path $Path1), $false, $true )
   $db2 = $dbe.OpenDatabase( (Resolve-Path $Path2), $false, $true )

   # Print table diffs (additional/missing table names).
   Get-Tables-Diff $db1 $db2
   
   # Table field differences.
                                    # {Exclude any system tables here.}
   Select-Common-Tables $db1 $db2 | Where-Object {-not $_.StartsWith("MSys")} | %{
      Get-Table-Fields-Diff $_ $db1 $db2
   }

   #--------------------------------------------------------------------------------------
   # Bulk of the scan.
                                    # {Exclude any system tables.}
   Select-Common-Tables $db1 $db2 | Where-Object {-not $_.StartsWith("MSys")} | %{
      $table = $_
   
      $fields = Select-Common-Fields $table $db1 $db2

      # Potentially unsafe tablename injection. Is there a better way?
      $fieldsquery = ($fields | %{"[$_]"}) -join ","
      $rs1 = $db1.OpenRecordset( "select $fieldsquery FROM $_" )
      $rs2 = $db2.OpenRecordset( "select $fieldsquery FROM $_" )

      $stopping = $false
      $differences = [RowDifferences]::new( $table, $fields )
      $stop_threshold = 3

      $row = 0

      while( $true ) {
         if( $rs1.EOF -or $rs2.EOF ) {
            if( (-not $rs1.EOF) -or (-not $rs2.EOF) ) {
               # We can't just check "RecordCount" because those do not contain the
               #  actual row count until you iterate/fetch all of the data. (I think?)
               Write-Output "(Data) Table $table has differing record counts."
            }
            break
         }

         # Not using RecordCount here because we aren't sure if that is guaranteed to be
         #  the current row.
         $row += 1
       
         # GetRows returns a 2 dimensional array. We'll just do one row at a time.
         $set1 = $rs1.GetRows(1)
         $set2 = $rs2.GetRows(1)

         # We only have a one dimension result, but make that official.
         $set1 = for( $i = 0; $i -lt $fields.Count; $i++ ) { $set1[$i,0] }
         $set2 = for( $i = 0; $i -lt $fields.Count; $i++ ) { $set2[$i,0] }
         
         for( $i = 0; $i -lt $fields.Count; $i++ ) {
            if( $set1[$i] -ne $set2[$i] ) {
               if( $differences.Rows1.Count -ge $stop_threshold ) {
                  $stopping = $true
                  $differences.Truncated = $true
                  break
               } else {
                  
                  $differences.Rows1 += [RowData]::new( $row, $set1 )
                  $differences.Rows2 += [RowData]::new( $row, $set2 )
                  break
               }
            }
         }
         if( $stopping ) { break }
      }
      
      $rs1.Close()
      $rs2.Close()

      if( $differences.Rows1.Count -gt 0 ) {
         $differences.Print()
      }
   }

   $db1.Close()
   $db2.Close()
}

#-----------------------------------------------------------------------------------------
class Match {
   [string]$MatchType = "Table Name"
}

#-----------------------------------------------------------------------------------------
class TableNameMatch : Match {
   [string]$Table

   TableNameMatch( [string]$Table ) {
      $this.MatchType = "Table Name"
      $this.Table = $Table
   }
}

#-----------------------------------------------------------------------------------------
class FieldNameMatch : Match {
   [string]$Table
   [string]$Field

   FieldNameMatch( [string]$Table, [string]$Field ) {
      $this.MatchType = "Field Name"
      $this.Table = $Table
      $this.Field = $Field
   }
}

#-----------------------------------------------------------------------------------------
# A bit of a fun exercise, these Match classes are defined in mdb-tools.ps1xml for
#  formatting output.
class FieldValueMatch : Match {
   [string]$Table
   [string]$Field
   [int]$Row
   [string]$Value
   
   FieldValueMatch( [string]$Table, [string]$Field, [int]$Row, [string]$Value ) {
      $this.MatchType = "Field Value"
      $this.Table = $Table
      $this.Field = $Field
      $this.Row   = $Row
      $this.Value = $Value
   }
}

#-----------------------------------------------------------------------------------------
Function Search-MDB {
   <#
      .SYNOPSIS
         Searches for a regex string in an MDB file.

      .PARAMETER Path
         Path to the MDB file to inspect.

      .PARAMETER SearchPattern
         Regex pattern to search against.

      .PARAMETER SearchInTableNames
         Default true; search table name strings for matches.

      .PARAMETER SearchInFieldNames
         Default true; search field name strings for matches.

      .PARAMETER SearchInFieldValues
         Default true; search ALL field/cell data values for matches.

      .EXAMPLE
         Search-MDB .\database1.mdb "test"

         Searches the database for the string "test". Will search in table names, 
   #>
   Param(
      [Parameter(Mandatory=$true)]
      [string]$Path,
      [Parameter(Mandatory=$true)]
      [string]$SearchPattern,
      [boolean]$SearchInTableNames  = $true,
      [boolean]$SearchInFieldNames  = $true,
      [boolean]$SearchInFieldValues = $true
   )

   $dbe = New-Object -comobject DAO.DBEngine.120
   $db = $dbe.OpenDatabase( (Resolve-Path $Path), $false, $true )

   foreach( $table in $db.TableDefs ) {
      # Skip internal tables.
      if( $table.Name.StartsWith("MSys") ) { continue }

      if( $SearchInTableNames -and ($table.Name -match $SearchPattern) ) {
         Write-Output( [TableNameMatch]::new($table.Name) )
      }
      
      $r = $db.OpenRecordset( "SELECT * FROM [$($table.Name)]" )
      
      if( $SearchInFieldNames ) {
         foreach( $field in $r.Fields ) {
            if( $field.Name -match $SearchPattern ) {
               Write-Output( [FieldNameMatch]::new($table.Name, $field.Name) )
            }
         }
      }
      
      if( $SearchInFieldValues ) {
         $rowCount = 0
         while( !$r.EOF ) {
            $rowCount++
            $rowData = $r.GetRows(1)
            for( $i = 0; $i -lt $r.Fields.Count; $i++ ) {
               if( [string]$rowData[$i,0] -match $SearchPattern ) {
                  $cellvalue = $rowData[$i,0]
                  Write-Output( [FieldValueMatch]::new(
                     $table.Name, $r.Fields[$i].Name, $rowCount, $cellvalue
                  ))
               }
            }
         }
      }

      $r.Close()
   }

   $db.Close()
}
#/////////////////////////////////////////////////////////////////////////////////////////