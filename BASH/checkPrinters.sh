  #!/bin/bash
  # This script will test the communication to each printer

  #for each printer the printers config file
  #filename="/home/ruimsousa/Documents/scripts/printers.conf"
  filename="/etc/cups/printers.conf"
  countPrinters=0

  # Declare arrays
  declare -a printerName
  declare -a printerIPaddr
  declare -a printerPing
  declare -a printerPort631
  declare -a printerPort9100

  if [[ -f "$filename" ]]; then
    echo "Reading file $filename"

    # read the file into array and count the number of printers
    readarray -t ARRAY <<< "$(sed -n -e '/^<Printer/p'  -e '/^Info/p' -e '/^DeviceURI/p' "$filename")"
    for line in "${!ARRAY[@]}"
    do
      #echo "${ARRAY[line]}"
      if [[ "${ARRAY[line]}" = \<Printer* ]]
      then
        ((countPrinters++))

      # get Printer Name
      temp1=${ARRAY[line+1]}
      printerName=(${printerName[@]} ${temp1#* })
      # get Printer IP Address
      temp1=${ARRAY[line+2]}
      temp1="${temp1#*//}"
      printerIPaddr=(${printerIPaddr[@]} ${temp1%%:*})
      #printerName[countPrinters-1]= ${p[1]}
      fi
    done

    # Print number of printer fond in file
    echo "Printers found: $countPrinters"
    echo ""

    #OUTPUT="#\tPRINTER_NAME\t\t\tIP_ADDRESS\tPING\tPORT_631\tPORT_9100\n\n"
    #OUTPUT="$OUTPUT----\t-----------------------------\t-------------------"
    #OUTPUT="$OUTPUT\t----------\t-------------\t-------------"

    # Start checking status of every printer
    for ((i = 0 ; i < $countPrinters ; i++ ));
    #for ((i = 0 ; i < 5 ; i++ ));
    do
      ##echo -ne "$i ${printerName[i]} ${printerIPaddr[i]}" | column -t;
      echo "Printer #$[i+1]"
      echo "Printer Name: ${printerName[i]}"
      echo "Printer IP: ${printerIPaddr[i]}"

      # test ping
      ping -c 1 ${printerIPaddr[i]} >/dev/null 2>&1
      if [ $? -ne 0 ] ; then #if ping exits nonzero...
        echo "PING: Failed"
        printerPing=(${printerPing[@]} "Failed")

        #echo "Port 631: Failed"
        #printerPort631=(${printerPort631[@]} "Failed")

        #echo "Port 9100: Failed"
        #printerPort9100=(${printerPort9100[@]} "Failed")

      else
        printerPing=(${printerPing[@]} "OK")
        echo "PING: OK"
      fi

      # test connection to socket 80
      nc -zv -w5 ${printerIPaddr[i]} 80 >/dev/null 2>&1
      if [ $? -ne 0 ] ; then #if nc exits nonzero...
        echo "Port 80: Failed"
        printerPort80=(${printerPort80[@]} "Failed")
      else
        echo "Port 80: OK"
        printerPort80=(${printerPort80[@]} "OK")
      fi

      # test connection to socket 631
      nc -zv -w5 ${printerIPaddr[i]} 631 >/dev/null 2>&1
      if [ $? -ne 0 ] ; then #if nc exits nonzero...
        echo "Port 631: Failed"
        printerPort631=(${printerPort631[@]} "Failed")
      else
        echo "Port 631: OK"
        printerPort631=(${printerPort631[@]} "OK")
      fi

      # test connection to socket 9100
      nc -zv -w5 ${printerIPaddr[i]} 9100 >/dev/null 2>&1
      if [ $? -ne 0 ] ; then #if nc exits nonzero...
        echo "Port 9100: Failed"
        printerPort9100=(${printerPort9100[@]} "Failed")
      else
        echo "Port 9100: OK"
        printerPort9100=(${printerPort9100[@]} "OK")
      fi


      #blank line
      echo ""

      #OUTPUT="$OUTPUT\n$[$i+1] ${printerName[i]} ${printerIPaddr[i]}"
      #OUTPUT="$OUTPUT ${printerPing[i]} ${printerPort631[i]}"
      #OUTPUT="$OUTPUT ${printerPort9100[i]}\n"

    done

    # Print output in tab-delimited column
    #echo -ne $OUTPUT | column -t
    header="\n%-4s %-20s %-15s %-8s %-6s %-10s %-10s\n"
    printf "$header" "#" "PRINTER NAME" "IP ADDRESS" "PING" "PORT 80" "PORT 631" "PORT 9100"
    printf "$header" "----" "--------------------" "---------------" "----------" "-------" "----------" "----------"
    format="%-4s %-20s %-15s %-8s %-6s %-10s %-10s\n"

    for ((i = 0 ; i < $countPrinters ; i++ ));
    do
      printf "$format" "$[$i+1]" "${printerName[i]}" "${printerIPaddr[i]}" "${printerPing[i]}" "${printerPort80[i]}" "${printerPort631[i]}" "${printerPort9100[i]}"
    done
    echo ""
    exit 0
  else
    echo "The specified file $filename doesn't exit."
    echo ""
    exit -1
  fi
