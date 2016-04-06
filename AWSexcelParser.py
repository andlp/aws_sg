import openpyxl, re, sys, os

from openpyxl import workbook,load_workbook
from subprocess import call # used to call the build_head.sh in the OS
filename = 'inputXLS1.xlsx'
wb = load_workbook(filename, use_iterators=True)
ws = wb.get_sheet_by_name('Sheet1') #ToDo: Logic to set this to arbitrary WS name ^vpc-*

#Get the temp files setup
head_out_file = open('build_head.sh', "w")
head_out_file.truncate() # clear the file contents

try:
	

	super_region = 'us-east-1'
	super_vpc_id = 'vpc-e940io34'
	for row in range(2, ws.max_row + 1): #Skip the first row of headers
		print ("\nWorking with Row # %s" % (row))
		
		#Set the Security Group Name
		super_sg_name = ws['A' + str(row)].value # take in the value from the spreadsheet
		match = re.match('[^\>]*', super_sg_name) #set the test flag anything that s not a < 
		if match:  #loop to tell the user
			print ('SG Name is good: %s' % (super_sg_name))
		else:
			print ("SG Name is NOT good: %s. Must be some letters or numbers" % (super_sg_name))
			break #ABORT if it fails 
		
		# Set the Security Group ID value
		super_sg_id = ws['B' + str(row)].value # take in the value from the spreadsheet
		match = re.match("^sg-[\\'a-zA-Z0-9\\']{8}", super_sg_id) #set the test flag
		if match:  #loop to tell the user
			print ('SG ID is good: %s' % (super_sg_id))
		else:
			print ("SG ID is NOT good: %s. Must be some letters or numbers" % (super_sg_id))
			break #ABORT if it fails
			
		# Set the VPC ID value
		super_vpc_id = ws['C' + str(row)].value # take in the value from the spreadsheet
		match = re.match("^vpc-[\\'a-zA-Z0-9\\']{8}", super_vpc_id) #set the test flag
		if match:  #loop to tell the user
			print ('VPC ID is good: %s' % (super_vpc_id))
		else:
			print ("VPC ID is NOT good: %s. Must be some letters or numbers" % (super_vpc_id))
			break #ABORT if it fails
		
		# Set the Description value	
		super_description = ws['D' + str(row)].value # take in the value from the spreadsheet
		match = re.match('[^\>]*', super_description) # any Char that isnt ">"
		if match:
			print ('Description is good: %s' % (super_description))
		else: 
			print ("Description is NOT good: %s.  Must be some symbol" % (super_description))
			break	
		
		# Set the Direction
		super_direction = ws['F' + str(row)].value # take in the value from the spreadsheet
		match = re.match('Ingress|Egress', super_direction) #set the test flag
		if match:  #loop to tell the user
			print ('Direction is good: %s' % (super_direction))
		else:
			print ("Direction is NOT good: %s. Must be Ingress or Egress" % (super_direction))
			break #ABORT if it fails
		
		# Set the Protocol
		super_proto = ws['G' + str(row)].value
		match = re.match('udp|tcp', super_proto)
		if match:
			print ('Protocol is good: %s' % (super_proto))
		else: 
			print ("Protocol is NOT good: %s.  Must be tcp or udp" % (super_proto))
			break			
	
		# Set the From_Port 
		super_from_port = ws['I' + str(row)].value
		super_from_port = str(super_from_port)
		match = re.match('[0-9]+', super_from_port)
		if match:
			print ("From Port is good: %s" % (super_from_port))
		else:
			print ("From Port is NOT good: %s.  Must be a number" % (super_from_port))
			break
		
		#Set the To_Port
		super_to_port = ws['H' + str(row)].value
		super_to_port = str(super_to_port) # digits have to be set to string for the re.match to work
		match = re.match('[0-9]+', super_to_port)
		if match:
			print ("To Port is good: %s" % (super_to_port))
		else:
			print ("To Port is NOT good: %s.  Must be a number" % (super_to_port))
			break
		
		# Set the CIDR IP
		super_cidr_ip = ws['J' + str(row)].value
		match = re.match('^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])(\/([0-9]|[1-2][0-9]|3[0-2]))$', super_cidr_ip )
		if match:
			print ("Cidr IP is good: %s" % (super_cidr_ip))
		else:
			break
		
		
		
		#Write the File Output
		if super_direction == "Ingress":
			head_out_file.write("aws ec2 authorize-security-group-ingress --group-name %s --protocol %s --port %s-%s --cidr %s;\n" %(super_sg_name, super_proto, super_from_port, super_to_port, super_cidr_ip))
		elif super_direction == "Egress":
			head_out_file.write("aws ec2 authorize-security-group-egress --group-id %s --ip-permissions '[{\"IpProtocol\": \"%s\", \"FromPort\": %s, \"ToPort\": %s, \"IpRanges\": [{\"CidrIp\": \"%s\"}]}]' \n"  %(super_sg_id, super_proto, super_from_port, super_to_port, super_cidr_ip))
	


		
finally:
    head_out_file.close()
    os.system("./build_head.sh")