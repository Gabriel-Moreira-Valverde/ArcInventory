import wmi

c = wmi.WMI()

# Fabricante e Modelo
system = c.Win32_ComputerSystem()[0]

# Sistema operacional
so = c.Win32_OperatingSystem()[0]
nome_so = so.Caption.replace("Microsoft ", "")

# Processador
p = c.Win32_Processor()[0]
nome_p = " ".join(p.Name.split()[0:3])

# Memória
memoria_bytes = sum(int(ram.Capacity) for ram in c.Win32_PhysicalMemory())
memoria_gb = memoria_bytes / (1024 ** 3)

# Disco
disco = c.Win32_DiskDrive()[0]
disco_modelo = disco.Model
disco_marca = disco_modelo.split()[0]
capacidade = int(disco.Size)
capacidade_gb = capacidade / (1024 ** 3)

disco_total = f"{disco_marca} {capacidade_gb:.2f} GB "

memoria_total = f"{memoria_gb:.0f} GB"