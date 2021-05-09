#-----------------------------------------------------------------------------------------
# MDB TOOLS
# Author: Mukunda Johnson
# License: MIT
#
# Handy scripts for inspecting MDB database files and differences.
#-----------------------------------------------------------------------------------------
Function Compare-MDBDatabase {
   <#
      .SYNOPSIS
         Compares two MS Access database files (MDB) and prints differences.

      .EXAMPLE
         Compare-MDBDatabase .\database1.mdb .\database2.mdb

         Compares the two database files given and prints differences to the output.
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

   #--------------------------------------------------------------------------------------
   # Outputs any difference in the table listing.
   Function Get-Tables-Diff {
      Param( $db1, $db2 )

      $tables1 = $db1.TableDefs | %{ $_.Name }
      $tables2 = $db2.TableDefs | %{ $_.Name }

      $tables1 | %{
         if( $tables2 -notcontains $_ ) {
               Write-Output "(Schema) DB1 has additional table $_"
         }
      }

      $tables2 | %{
         if( $tables1 -notcontains $_ ) {
               Write-Output "(Schema) DB2 has additional table $_"
         }
      }
   }

   #--------------------------------------------------------------------------------------
   # Intersects the table list and returns a list of names that are present in both
   #  tables.
   function Select-Common-Tables {
      Param( $db1, $db2 )
      $tables1 = $db1.TableDefs | %{ $_.Name }
      $tables2 = $db2.TableDefs | %{ $_.Name }
      
      $tables1 | Where-Object {$tables2 -contains $_}
   }

   Get-Tables-Diff $db1 $db2

   #--------------------------------------------------------------------------------------
   # Checks field list in both tables and outputs any differences.
   function Get-Table-Fields-Diff {
      Param( $name, $db1, $db2 )
      $fields1 = $db1.TableDefs | Where-Object {$_.Name -eq $name}
      $fields2 = $db2.TableDefs | Where-Object {$_.Name -eq $name}

      $not_found = [System.Collections.ArrayList]@(($fields2 | %{ $_.Name }))

      foreach( $ca in $fields1 ) {
         $fieldname = $ca.Name
         $cb = $fields2 | Where-Object {$_.Name -eq $fieldname}
         if( -not $cb ) {
            Write-Output "(Schema) DB1 has additional field in ${name}: $fieldname"
            continue
         }

         $not_found.Remove( $fieldname )

         $comparisons = ("Type", "Size", "DefaultValue", "Required", "AllowZeroLength")

         $comparisons | %{
            if( $ca.$_ -ne $cb.$_ ) {
               Write-Output "$name.$fieldname - `"$_`" differs: $($ca.$_) - $($cb.$_)"
            }
         }
      }
      
      foreach( $a in $not_found ) {
         Write-Output "(Schema) DB2 has additional field in ${name}: $a"
      }
   }

   #--------------------------------------------------------------------------------------
   # Returns a list of field names that are present in both tables.
   function Select-Common-Fields {
      Param( $table, $db1, $db2 )
      $fields1 = @(($db1.TableDefs | Where-Object {$_.Name -eq $table}).Fields | %{$_.Name})
      $fields2 = @(($db2.TableDefs | Where-Object {$_.Name -eq $table}).Fields | %{$_.Name})

      $fields1 | Where-Object {$fields2 -contains $_}
   }

   #--------------------------------------------------------------------------------------
   # Bulk of the scan.
   #                                # Exclude any system tables.
   Select-Common-Tables $db1 $db2 | Where-Object {-not $_.StartsWith("MSys")} | %{
      $table = $_
      Get-Table-Fields-Diff $table $db1 $db2

      $fields = Select-Common-Fields $table $db1 $db2

      # Potentially unsafe table injection method. Is there a better way?
      $fieldsquery = ($fields | %{"[$_]"}) -join ","
      $rs1 = $db1.OpenRecordset( "select $fieldsquery FROM $_" )
      $rs2 = $db2.OpenRecordset( "select $fieldsquery FROM $_" )

      $differing = New-Object System.Collections.ArrayList
      $stop_threshold = 3
      $stopping = $false

      $row = 0

      while( $true ) {
         if( $rs1.EOF -or $rs2.EOF ) {
            if( -not $rs.EOF -or -not $rs2.EOF ) {
               # We can't just check "RecordCount" because those do not contain the
               #  actual row count until you iterate/fetch all of the data. (I think?)
               Write-Output "(Data) Table $_ has differing record counts."
            }
            break
         }

         $row += 1
         # GetRows returns a 2 dimensional array. We'll just do one row at a time.
         $set1 = $rs1.GetRows(1)
         $set2 = $rs2.GetRows(1)

         for( $i = 0; $i -lt $fields.count; $i++ ) {
            if( $set1[$i,0] -ne $set1[$i,0] ) {
               if( $differing.Count -ge $stop_threshold ) {
                  $stopping = $true
                  break
               } else {
                  $ht = @{ Row = "$row(1)" }
                  for( $j = 0; $j -lt $fields.count; $j++ ) {
                     $a = $set1[$j, 0]
                     $ht.Add( $fields[$j], $a )
                  }
                  $differing.Add( $ht )
                  $ht = @{ Row = "$row(2)" }
                  for( $j = 0; $j -lt $fields.count; $j++ ) {
                     $a = $set2[$j, 0]
                     $ht.Add( $fields[$j], $a )
                  }
                  $differing.Add( $ht )
               }
            }
         }
         if( $stopping ) { break }
      }

      if( $differing.Count -gt 0 ) {
         $differing | Format-Table
      }

      if( $stopping ) {
         Write-Output "More than 3 differences; skipping to next table."
      }

      $rs1.Close()
      $rs2.Close()
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
Function Search-MDBDatabase {
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
         Search-MDBDatabase .\database1.mdb "test"

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