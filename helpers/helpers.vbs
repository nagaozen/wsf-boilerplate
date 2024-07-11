' File: helpers.vbs
' 
' Copyright 2017 Fabio Zendhi Nagao <nagaozen[at]gmail.com>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of
' this software and associated documentation files (the “Software”), to deal in
' the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
' the Software, and to permit persons to whom the Software is furnished to do so,
' subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
' CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

function bin_to_base64 (bin)
  with createObject("MSXML2.DOMDocument.6.0")
    with .createElement("data")
      .dataType = "bin.base64"
      .nodeTypedValue = bin
      bin_to_base64 = replace(.text, vbLF, "")
    end with
  end with
end function

function base64_to_bin (base64)
  with createObject("MSXML2.DOMDocument.6.0")
    with .createElement("data")
      .dataType = "bin.base64"
      .text = base64
      base64_to_bin = .nodeTypedValue
    end with
  end with
end function

function environment_variable (byVal name)
  with createObject("WScript.Shell")
    environment_variable = .expandEnvironmentStrings(name)
  end with
end function

function date_to_iso8601Z (date)
  dim zulu, dt, tm
  with createObject("WScript.Shell")
    zulu = dateadd("n", .RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"), date)
  end with
  dt = join(array(_
    datepart("yyyy", zulu), _
    right("0" & datepart("m", zulu), 2), _
    right("0" & datepart("d", zulu), 2) _
  ), "")
  tm = join(array(_
    right("0" & datepart("h", zulu), 2), _
    right("0" & datepart("n", zulu), 2), _
    right("0" & datepart("s", zulu), 2) _
  ), "")
  date_to_iso8601Z = join( array(dt,tm), "T" ) & "Z"
end function

class RuntimeError
  public number
  public source
  public description
  public default function [new](number, source, description)
    Me.number = number
    Me.source = source
    Me.description = description
    set [new] = Me
  end function
end class

function parse_json (value)
  dim exception
on error resume next : Err.clear
  set safe_datasource = JSON.parse(value)
  if Err.number <> 0 then
    set exception = (new RuntimeError)(13, "JSON.parse", "User provided value `" & value & "` MUST conform to JSON specification.")
  end if
on error goto 0
  if not isEmpty(exception) then
    parse_json = array(exception, empty)
  else
    parse_json = array(empty, safe_datasource)
  end if
end function

function fetch (url, sk, verb, payload)
  const resolveTimeout = 60000
  const connectTimeout = 60000
  const sendTimeout    = 120000
  const receiveTimeout = 240000
  with createObject("MSXML2.ServerXMLHTTP.6.0")
    .setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
    .open verb, url, false
    .setRequestHeader "Ocp-Apim-Subscription-Key", sk
    .setRequestHeader "Content-Type", "application/json"
    .send payload
    fetch = .responseText
  end with
end function

function file_exists (filespec)
  with createObject("Scripting.FileSystemObject")
    file_exists = .fileExists(filespec)
  end with
end function

function get_file_extension (filespec)
  with createObject("Scripting.FileSystemObject")
    get_file_extension = .getExtensionName(filespec)
  end with
end function

function read_utf8_file (filespec)
  if not file_exists(filespec) then Err.raise 53, "file_exists runtime error", "File not found. <" & filespec & ">"
  with createObject("ADODB.Stream")
    .type = adTypeText
    .charset = "UTF-8"
    .mode = adModeReadWrite
    .open
    .loadFromFile filespec
    read_utf8_file = .readText(adReadAll)
    .close
  end with
end function

sub save_utf8_file (filespec, text)
  with createObject("ADODB.Stream")
    .type = adTypeText
    .charset = "UTF-8"
    .mode = adModeReadWrite
    .open
    .writeText text
    .setEOS
    .position = 0
    .saveToFile filespec, adSaveCreateOverwrite
    .close
  end with
end sub

function read_bin_file (filespec)
  if not file_exists(filespec) then Err.raise 53, "file_exists runtime error", "File not found. <" & filespec & ">"
  with createObject("ADODB.Stream")
    .type = adTypeBinary
    .mode = adModeReadWrite
    .open
    .loadFromFile filespec
    read_bin_file = .read(adReadAll)
    .close
  end with
end function

sub save_bin_file (filespec, buffer)
  with createObject("ADODB.Stream")
    .type = adTypeBinary
    .mode = adModeReadWrite
    .open
    .write buffer
    .setEOS
    .position = 0
    .saveToFile filespec, adSaveCreateOverwrite
    .close
  end with
end sub

function compare_docx (byVal lhs_filespec, byVal rhs_filespec)
  dim lhs, rhs, _
      diff, revision, changes, change

  with createObject("Scripting.FileSystemObject")
    lhs_filespec = .getAbsolutePathName(lhs_filespec)
    rhs_filespec = .getAbsolutePathName(rhs_filespec)
  end with

  ' [Word VBA reference](https://learn.microsoft.com/en-us/office/vba/api/overview/word)
  with createObject("Word.Application")
    .visible = false

    set lhs = .Documents.open(lhs_filespec)
    set rhs = .Documents.open(rhs_filespec)

    set changes = JSON.parse("[]")
    set diff = .Application.compareDocuments(lhs, rhs)
    for each revision in diff.Revisions
      set change = JSON.parse("{}")
      ' [Word.Revision](https://learn.microsoft.com/en-us/office/vba/api/word.revision#properties)
      change.set "type", revision.type
      change.set "text", revision.range.text
      change.set "author", revision.author
      change.set "date", revision.date
      changes.push change
      set change = nothing
    next
    set compare_docx = changes
    diff.close false
    set diff = nothing
    set changes = nothing

    rhs.close false
    lhs.close false
    set rhs = nothing
    set lhs = nothing

    .quit false
  end with
end function

sub docx_to_pdf (byVal input_filespec, byVal output_filespec)
  dim document

  with createObject("Scripting.FileSystemObject")
    input_filespec = .getAbsolutePathName(input_filespec)
    output_filespec = .getAbsolutePathName(output_filespec)
  end with

  with createObject("Word.Application")
    .visible = false

    set document = .Documents.open(input_filespec)
    document.saveAs2 output_filespec, 17' wdFormatPDF <https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat>
    document.close false
    set document = nothing

    .quit false
  end with
end sub

sub docx_to_utf8 (byVal input_filespec, byVal output_filespec)
  dim document

  with createObject("Scripting.FileSystemObject")
    input_filespec = .getAbsolutePathName(input_filespec)
    output_filespec = .getAbsolutePathName(output_filespec)
  end with

  with createObject("Word.Application")
    .visible = false

    set document = .Documents.open(input_filespec)
    ' https://learn.microsoft.com/en-us/office/vba/api/word.saveas2
    document.saveAs2 output_filespec, 7, false, , , , , , , , , 65001
    document.close false
    set document = nothing

    .quit false
  end with
end sub
