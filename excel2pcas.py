#-------------------------------------------------------------------------------
# Name:        excel2pcas
# Purpose:     Convert an Excel file to a PCAS one, via intermediate conversion 
#              to a NEXUSC Interchange File
#
# Author:      Jesus Ruiz
#
# Created:     20/06/2014
# Licence:     Apache 2.0
#-------------------------------------------------------------------------------

import csv
import os.path
import datetime
import argparse
import requests
import urlparse
import zipfile
import os
import time
import calendar
import math
import json
from StringIO import StringIO

# The structure of the NEXUSC Header record
NEXUSC_header_layout = [
    ( 'code', 0, 1 ), # Must be 'C' for a header record
    ( 'university', 1, 3 ),
    ( 'creation_date', 4, 10 ),
    ( 'sequence', 14, 1 ),
    ( 'creation_date_prev', 15, 10 ),
    ( 'sequence_prev', 25, 1 ),
    ( 'filler1', 26, 335 ),
    ( 'EOL', 361, 1 ),
]
NEXUSC_headerlen= 362

# The structure of the NEXUSC Data record
NEXUSC_record_layout = [
    ( 'code', 0, 1 ), # Must be 'D' for a data record
    ( 'doc_type', 1, 1 ),
    ( 'doc_number', 2, 9 ),
    ( 'name', 11, 20 ),
    ( 'surname1', 31, 26 ),
    ( 'surname2', 57, 26 ),
    ( 'domicilio', 83, 103 ),
    ( 'studentID', 186, 12 ),
    ( 'filler1', 198, 155 ),
    ( 'birthdate', 353, 8 ),
    ( 'EOL', 361, 1 ),
]
NEXUSC_recordlen= 362

# The structure of the NEXUSC Footer record
NEXUSC_footer_layout = [
    ( 'code', 0, 1 ), # Must be 'F' for the footer record
    ( 'university', 1, 3 ),
    ( 'num_records', 4, 5 ), # The number of data records that should be in the file
    ( 'filler1', 9, 352 ),
    ( 'EOL', 361, 1),
]
NEXUSC_footerlen= 362

# -------------------------------------------------------
# Define some utility functions

# Function to convert a long integer to a PIC S9(10)V COMP-3 field from COBOL
def longToCOMP3( v, digits ):
    num_bytes = int(math.ceil((digits+1.0)/2))
    num_chars = 2*num_bytes
    b = bytearray(num_bytes)
    vs = format( v, "020") + chr(ord("0")+0xC)
    vs = vs[-num_chars:]
    for i in range(num_bytes):
        nibble_high = ord(vs[2*i]) - ord("0")
        nibble_low = ord(vs[2*i+1]) - ord("0")
        b[i] = (nibble_high << 4) + nibble_low
    return b

def initlog():
    with open( "ntopcas.log", 'w' ) as f:
        f.write( datetime.datetime.today().isoformat() + " - Initialized\n" )

def dolog(text):
    with open( "ntopcas.log", 'a' ) as f:
        f.write( datetime.datetime.today().isoformat() + " - " + text + "\n" )


# -------------------------------------------------------
# Execution starts here

def main():
    
    initlog()
    
    parser = argparse.ArgumentParser(description='Create PCAS file from Excel file via NEXUSC Interchange files.')
    parser.add_argument("fileName", help="The EXCEL file name with student data")

    args = parser.parse_args()

    # -------------------------------------------------------

    # Get the EXCEL file name rom the command line
    excelFileName = args.fileName
    
    # Generate the NEXUSC interchange file name
    basename, extension = os.path.splitext(os.path.basename(excelFileName))
    interFileName = basename + ".TXT"

    # Check if the file exists
    if os.path.isfile( excelFileName ) == False:
        print "File " + excelFileName + " does not exist"
        return

    # Process the EXCEL file
    with open(excelFileName, 'rb') as csvfile:
        dialect = csv.Sniffer().sniff(csvfile.read(1024))
        csvfile.seek(0)
#        reader = csv.DictReader(csvfile, dialect=dialect)
        reader = csv.reader(csvfile, dialect=dialect)
        
        with open( interFileName, 'wb' ) as nfile:

            # -------------------------------------------------------
            # Generate the header record for NEXUSC
            nfile.write( "C" )
            
            # Write the university code for the test university
            nfile.write( "111" )
            
            # Write today as the date of creation of the file
            nfile.write( (datetime.date.today()).strftime("%Y%m%d") )
            
            # Pad with 2 blank characters
            nfile.write( "  " )
            
            # Write the sequence, which will be always "0"
            nfile.write( "0" )
            
            # Fake date of creation of previous file: same as today 
            nfile.write( (datetime.date.today()).strftime("%Y%m%d") )

            # Pad with 2 blank characters
            nfile.write( "  " )
            
            # Write the sequence, which will be always "0"
            nfile.write( "0" )
            
            # Write the filler
            nfile.write( "".ljust( 335 ) )

            # Write the EOL character
            nfile.write ( "\n" )

    
            # Iterate the rows in the EXCEL file to write the data records
            line = 0
            for row in reader:
                line = line + 1
                # Discard the first line
                if line > 1:
                    
                    # Write the indicator for the data record
                    nfile.write( "D" )
                    
                    # Write filler
                    nfile.write( "".ljust( 10 ) )
                    
                    # Write the Name of the student
                    nfile.write( row[1].ljust( 20 ) )
                    
                    # Write the Surname of the student
                    nfile.write( row[2].ljust( 26 ) )
                    
                    # Write filler
                    nfile.write( "".ljust( 129 ) )
                    
                    # Write the StudentID
                    nfile.write( row[0].ljust( 12 ) )
                    
                    # Write filler
                    nfile.write( "".ljust( 75 ) )
                    
                    # Write the Title
                    nfile.write( row[4].ljust( 20 ) )
                    
                    # Write the Department
                    nfile.write( row[5].ljust( 30 ) )

                    # Write the Position
                    nfile.write( row[6].ljust( 30 ) )
                    
                    # Write the Birthdate
                    nfile.write( row[3].ljust( 8 ) )

                    # Write the EOL character
                    nfile.write ( "\n" )
                    
            # -------------------------------------------------------
            # Generate the footer record
            nfile.write( "F" )
            
            # Write the university code for the test university
            nfile.write( "111" )
            
            # Write the number of records written
            nfile.write( str(line-1).zfill(5))
            
            # Write the filler
            nfile.write( "".ljust( 352 ) )

            # Write the EOL character
            nfile.write ( "\n" )            

    # Print a success message
    print "NEXUSC file '" + interFileName + "' succesfully created,", (line -1), "records written."

    # Build the PCAS file name
    pcasFileName = basename + ".PCAS.TXT"        
    
    # Initialize the variable
    NEXUSC_records = []

    # Open the NEXUSC interchange file for reading in binary mode
    with open( interFileName, 'rb' ) as interfile:
        
        # Read in memory the whole NEXUSC Interchange file
        # And store the Data records in the list NEXUSC_records

        for line in interfile:

            # Check if Header record
            if line[0] == 'C':

                NEXUSC_header = dict()
                for name, start, size in NEXUSC_header_layout:
                    # Check if line has enough length for this field
                    if start + size > len( line ):
                        break
                    NEXUSC_header[name] = line[start:start+size]

            # Check if Data record
            elif line[0] == 'D':

                data_record = dict()
                for name, start, size in NEXUSC_record_layout:
                    # Check if line has enough length for this field
                    if start + size > len( line ):
                        break
                    data_record[name]= line[start:start+size]
                NEXUSC_records.append(data_record)

            # Check if Footer record
            elif line[0] == 'F':

                NEXUSC_footer = dict()
                for name, start, size in NEXUSC_footer_layout:
                    # Check if line has enough length for this field
                    if start + size > len( line ):
                        break
                    NEXUSC_footer[name] = line[start:start+size]

            else:
                print "Error: record type is wrong!!: " + line[0]

    # Check if the number of records actually read is the same as the one specified in the Footer record
    if len(NEXUSC_records) != int( NEXUSC_footer["num_records"] ) :
        print "Error: num_records field in Footer is " + NEXUSC_footer["num_records"] + " but file has", len(NEXUSC_records)
        return
    else:
        print "File", interFileName, "processed successfully"
        print "The number of Data records processed is:", len(NEXUSC_records)


    # -------------------------------------------------------
    # Create the PCAS file and overwrite any file with the same name
    
    with open( pcasFileName, 'wb' ) as pcasfile:

        # -------------------------------------------------------
        # Generate the header record for PCAS
        pcasfile.write( "3294" ) # MPINTCAB-CODENT

        # Generate the Unique identifier for the file
        dani_ts = calendar.timegm(time.strptime('Wed Jul 23 18:00:00 2008'))
        counter = long(time.time() - dani_ts)
        # And write it in a PIC S9(10)V COMP-3
        uniqueFileId_COMP3 =  longToCOMP3( counter, 10 )
        pcasfile.write( uniqueFileId_COMP3 ) # MPINTCAB-NSECFIC

        # Write the number 1 in a PIC S9(02)V COMP-3
        pcasfile.write( longToCOMP3( 1, 2 ) )

        pcasfile.write( "C" )

        d = NEXUSC_header["creation_date"]
        creation_date = d[0:4]+"-"+d[4:6]+"-"+d[6:8]
        pcasfile.write( creation_date )
        pcasfile.write( time.strftime("%H:%M:%S") )

        # Write the number of records, including the header
        # The format in the file is a PIC S9(12)V COMP-3
        pcasfile.write( longToCOMP3( len(NEXUSC_records)+ 1, 12 ) )

        # Write blanks as filler up to the record length (1000 bytes)
        pcasfile.write( "".ljust( 962 ) )

        # -------------------------------------------------------
        # Write the data records
        for r in NEXUSC_records:

            pcasfile.write( "3294" ) # MPINTCAB-CODENT

            # Write the unique file ID in a PIC S9(10)V COMP-3
            pcasfile.write( uniqueFileId_COMP3 ) # MPINTCAB-NSECFIC

            # Write the number 1 in a PIC S9(02)V COMP-3
            pcasfile.write( longToCOMP3( 1, 2 ) )

            # Write a "D" to indicate a Data record
            pcasfile.write( "D" )

            # Write the constant "3294"
            pcasfile.write( "3294" ) # MPINTCAB-CODENT

            # University ID padded with blanks to the right
            university = NEXUSC_header["university"].ljust( 4 )
            pcasfile.write( university )

            # Student ID padded with blanks to the right
            studentID = r["studentID"].ljust( 20 )
            pcasfile.write( studentID )

            # The first 15 chars of First name of student
            name = r["name"][0:15]
            pcasfile.write( name )

            # The first 20 chars of Surname of student
            surname1 = r["surname1"][0:20]
            pcasfile.write( surname1 )

            # Birthdate of the student (YYYYMMDD), padded with blanks to the right
            d = r["birthdate"]
            birthdate = d[0:4]+"-"+d[4:6]+"-"+d[6:8]
            pcasfile.write( birthdate )

            # Write blanks as filler up to the record length (1000 bytes)
            pcasfile.write( "".ljust( 914 ) )

    # Print a success message
    print "PCAS file '" + pcasFileName + "' succesfully created,", len(NEXUSC_records), "records written"

if __name__ == '__main__':
    main()

