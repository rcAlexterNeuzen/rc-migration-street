Function Install-Requirements {
    param (
        [Parameter(Mandatory)]
        [array]$modules,
        [Parameter(Mandatory)]
        [string]$file
    )
    #Check for required Modules and versions; Prompt for install if missing and import.
    $installedModules = Get-InstalledModule

    ForEach ($Module in $modules) {
        if ($installedModules.name -contains $module) {
            $module = $installedModules | Where-Object { $_.Name -eq $module }
            Log-Message -file $file -Status DONE -Message "$($module.name) is installed with version $($module.version)."
            Write-Host "    Checking for latest version ..." -NoNewLine
            Log-Message -file $file -Status INFO -Message "    Checking for latest version ..."
            $latest = Find-Module -Name $module -Repository PSGallery -ErrorAction Stop
            if ($Latest.version -gt $module.Version) {
                Log-Message -file $file -Status INFO -Message "Module needs to be updated to $($Latest.version)"
            }
            else {
                Log-Message -file $file -Status INFO -Message "Module is latest version"
            }
        }
        else {
            Log-Message -file $file -Status WARNING -Message "$module is not installed"
            Log-Message -file $file -Status INFO -Message "    Checking for latest version ..."
            $latest = Find-Module -Name $module -Repository PSGallery -ErrorAction Stop
            if ($Latest.version -gt $Module.Version) {
                Write-host "$($Latest.version)" -ForegroundColor Yellow
                Write-Host "    Installing latest version ..." -NoNewLine -ForegroundColor Yellow
                
                Try {
                    Install-Module $module -RequiredVersion $latest.version -Repository PSGallery -Force 
                    Log-Message -file $file -Status INFO -Message "Installing latest version ..."
                }
                catch {
                    Log-Message -file $file -Status error -Message "Installing latest version ..."
                }

            }
            else {
                Log-Message -file $file -Status done -Message "$Module is latest version"
            }
        }
    }
}

function ConvertTo-Psd {
	[OutputType([String])]
	param(
		[Parameter(Position=0, ValueFromPipeline=1)]
		$InputObject,
		[int] $Depth,
		[string] $Indent
	)
	begin {
		$objects = [System.Collections.Generic.List[object]]@()
	}
	process {
		$objects.Add($InputObject)
	}
	end {
		trap {Throw $_}

		$script:Depth = $Depth
		$script:Pruned = 0
		$script:Indent = Convert-Indent $Indent
		$script:Writer = New-Object System.IO.StringWriter
		try {
			foreach($object in $objects) {
				Write-Psd $object
			}
			$script:Writer.ToString().TrimEnd()
			if ($script:Pruned) {Write-Warning "ConvertTo-Psd truncated $script:Pruned objects."}
		}
		finally {
			$script:Writer = $null
		}
	}
}

function Convert-Indent($Indent) {
	switch($Indent) {
		'' {return '    '}
		'1' {return "`t"}
		'2' {return '  '}
		'4' {return '    '}
		'0' {return ''}
	}
	$Indent
}

function Write-Psd($Object, $Depth=0, [switch]$NoIndent) {
	$indent1 = $script:Indent * $Depth
	if (!$NoIndent) {
		$script:Writer.Write($indent1)
	}

	if ($null -eq $Object) {
		$script:Writer.WriteLine('$null')
		return
	}

	$type = $Object.GetType()
	switch([System.Type]::GetTypeCode($type)) {
		Object {
			if ($type -eq [System.Guid] -or $type -eq [System.Version]) {
				$script:Writer.WriteLine("'{0}'", $Object)
				return
			}
			if ($type -eq [System.Management.Automation.SwitchParameter]) {
				$script:Writer.WriteLine($(if ($Object) {'$true'} else {'$false'}))
				return
			}
			if ($type -eq [System.Uri]) {
				$script:Writer.WriteLine("'{0}'", $Object.ToString().Replace("'", "''"))
				return
			}
			if ($script:Depth -and $Depth -ge $script:Depth) {
				$script:Writer.WriteLine("''''")
				++$script:Pruned
				return
			}
			if ($Object -is [System.Collections.IDictionary]) {
				if ($Object.Count) {
					$itemNo = 0
					$script:Writer.WriteLine('@{')
					$indent2 = $script:Indent * ($Depth + 1)
					foreach($e in $Object.GetEnumerator()) {
						$key = $e.Key
						$value = $e.Value
						$keyType = $key.GetType()
						if ($keyType -eq [string]) {
							if ($key -match '^\w+$' -and $key -match '^\D') {
								$script:Writer.Write('{0}{1} = ', $indent2, $key)
							}
							else {
								$script:Writer.Write("{0}'{1}' = ", $indent2, $key.Replace("'", "''"))
							}
						}
						elseif ($keyType -eq [int]) {
							$script:Writer.Write('{0}{1} = ', $indent2, $key)
						}
						elseif ($keyType -eq [long]) {
							$script:Writer.Write('{0}{1}L = ', $indent2, $key)
						}
						elseif ($script:Depth) {
							++$script:Pruned
							$script:Writer.Write('{0}item__{1} = ', $indent2, ++$itemNo)
							$value = New-Object 'System.Collections.Generic.KeyValuePair[object, object]' $key, $value
						}
						else {
							throw "Not supported key type '$($keyType.FullName)'."
						}
						Write-Psd $value ($Depth + 1) -NoIndent
					}
					$script:Writer.WriteLine("$indent1}")
				}
				else {
					$script:Writer.WriteLine('@{}')
				}
				return
			}
			if ($Object -is [System.Collections.IEnumerable]) {
				$script:Writer.Write('@(')
				$empty = $true
				foreach($e in $Object) {
					if ($empty) {
						$empty = $false
						$script:Writer.WriteLine()
					}
					Write-Psd $e ($Depth + 1)
				}
				if ($empty) {
					$script:Writer.WriteLine(')')
				}
				else {
					$script:Writer.WriteLine("$indent1)" )
				}
				return
			}
			if ($Object -is [scriptblock]) {
				$script:Writer.WriteLine('{{{0}}}', $Object)
				return
			}
			if ($Object -is [PSCustomObject] -or $script:Depth) {
				$script:Writer.WriteLine('@{')
				$indent2 = $script:Indent * ($Depth + 1)
				foreach($e in $Object.PSObject.Properties) {
					$key = $e.Name
					if ($key -match '^\w+$' -and $key -match '^\D') {
						$script:Writer.Write('{0}{1} = ', $indent2, $key)
					}
					else {
						$script:Writer.Write("{0}'{1}' = ", $indent2, $key.Replace("'", "''"))
					}
					Write-Psd $e.Value ($Depth + 1) -NoIndent
				}
				$script:Writer.WriteLine("$indent1}" )
				return
			}
		}
		String {
			$script:Writer.WriteLine("'{0}'", $Object.Replace("'", "''"))
			return
		}
		Boolean {
			$script:Writer.WriteLine($(if ($Object) {'$true'} else {'$false'}))
			return
		}
		DateTime {
			$script:Writer.WriteLine("[DateTime] '{0}'", $Object.ToString('o'))
			return
		}
		Char {
			$script:Writer.WriteLine("'{0}'", $Object.Replace("'", "''"))
			return
		}
		DBNull {
			$script:Writer.WriteLine('$null')
			return
		}
		default {
			if ($type.IsEnum) {
				$script:Writer.WriteLine("'{0}'", $Object)
			}
			else {
				$script:Writer.WriteLine($Object)
			}
			return
		}
	}

	throw "Not supported type '{0}'." -f $type.FullName
}