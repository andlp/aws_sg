echo "starting on line 1"
aws ec2 authorize-security-group-ingress --group-name DEV-SPSS-W1-PRIV-APP --protocol tcp --port 121-130 --cidr 10.201.180.105/32;
aws ec2 authorize-security-group-egress --group-id sg-4d3f1d35 --ip-permissions '[{"IpProtocol": "tcp", "FromPort": 1433, "ToPort": 1533, "IpRanges": [{"CidrIp": "10.201.180.105/32"}]}]' 
aws ec2 authorize-security-group-egress --group-id sg-4d3f1d35 --ip-permissions '[{"IpProtocol": "udp", "FromPort": 1434, "ToPort": 1434, "IpRanges": [{"CidrIp": "10.201.180.105/32"}]}]' 
aws ec2 authorize-security-group-ingress --group-name PROD-SPSS-W1-PRIV-DB --protocol tcp --port 1433-1433 --cidr 10.201.180.105/32;
aws ec2 authorize-security-group-ingress --group-name PROD-SPSS-W1-PRIV-DB --protocol udp --port 1434-1434 --cidr 10.201.180.105/32;
