# ipfunctions
Excel IP Functions Addon

If you use Excel for doing much of anything with IP addresses or router configs, then you know there are a lot of limitations when working with IP's and subnets. This is an Excel AddIn I have been working on to help solve that problem. It adds a bunch of functions that allow you to much more easily work with and manipulate IP addresses.
Adds an IP Functions group of functions selectable from the Insert Functions button. Functions include:

IPOCTET - returns the given octet
IPNetwork - returns the network
IPHOSTS - returns the number of hosts in a given subnet
IPMASKVAL - returns the mask from the mask length
IPMASKWILD - returns the wildcard mask
IPNEXTNET - returns the next subnet of the same size
Many more…

Adds a few buttons to the Data ribbon toolbar for:
IPSort - sorts a selection numerically by IP Address
Fill IP Address - Fills a selection with sequential IP addresses
Fill Subnets - Fills a selection in progressive subnets of the same size
Convert - Removes formulas and converts cell contents to text

The majority of the functions use IP addresses in the form a.b.c.d or a.b.c.d/x. 

