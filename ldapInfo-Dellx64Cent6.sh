#! /bin/bash

#printf "\n\nHostname: /etc/sysconfig/network\n"
#cat /etc/sysconfig/network | grep -i hostname;
printf "\nFiles:"
printf '\t\t%s\n' "/etc/sysconfig/network" "/etc/sysconfig/network-scripts/ifcfg-eth0" "/etc/resolv.conf" "/etc/hosts" "/etc/hosts.bak"

#printf "\n/etc/sysconfig/network"
printf "\nHostname:\t" && hostname -s
printf "Domain:\t\t";dnsdomainname
#printf "\n/etc/sysconfig/network-scripts/ifcfg-eth0"
printf "IP:\t\t"; hostname -i

#printf "\n/etc/resolv.conf"
printf "Resolvers:" && printf '\t%s\n' "$(cut -d ' ' -f 2  /etc/resolv.conf | sed '1q;d')"
printf '\t\t%s\n' "$(cut -d ' ' -f 2 /etc/resolv.conf | sed '2q;d')"

printf '\n%s' "$(cat /etc/hosts)" "$(cat /etc/hosts.bak | sed '2q;d')"

printf "\nNIS Domain:\t"; nisdomainname
printf "YP Domain:\t";  ypdomainname
printf "Users: \t"; ls /home

printf "\nOpenLDAP Conf: /etc/openldap/slapd.d/cn=config/olcDatabase={2}bdb.ldif\n"
cat /etc/openldap/slapd.d/cn\=config/olcDatabase={2}bdb.ldif | grep olcRootPW:
cat /etc/openldap/slapd.d/cn\=config/olcDatabase={2}bdb.ldif | grep olcSuffix:
cat /etc/openldap/slapd.d/cn\=config/olcDatabase={2}bdb.ldif | grep olcRootDN:

printf "\nOpenLDAP RootDN: /etc/openldap/slapd.d/cn=config/olcDatabase={1}monitor\n"
cat /etc/openldap/slapd.d/cn\=config/olcDatabase={1}monitor.ldif | grep olcAccess:

printf "\nOpenLDAP Client Conf: /etc/openldap/ldap.conf\n"
cat /etc/openldap/ldap.conf | grep BASE
cat /etc/openldap/ldap.conf | grep URI

printf "\nOpenLDAP DB Search Results: grep #\n";
printf "\n"; ldapsearch -D "cn=Manager,dc=lpic,dc=local" -W '(objectclass=*)' | grep \#

printf "\n";printf "\n"