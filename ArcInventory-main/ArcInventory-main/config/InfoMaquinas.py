import wmi

c = wmi.WMI()

# Tipo da maquina
chassis = c.Win32_SystemEnclosure()[0].ChassisTypes

notebook_types = {8, 9, 10, 14}

if any(t in notebook_types for t in chassis):
    tipo_pc = "Notebook"
else:
    tipo_pc = "Desktop"

# Fabricante e Modelo
system = c.Win32_ComputerSystem()[0]

# Sistema operacional
so = c.Win32_OperatingSystem()[0]
nome_so = so.Caption.replace("Microsoft ", "")

# Processador
p = c.Win32_Processor()[0]
nome_p = " ".join(p.Name.split()[0:3])

# Memória
memorias = c.Win32_PhysicalMemory()

memoria_bytes = sum(int(ram.Capacity) for ram in memorias)
memoria_gb = memoria_bytes / (1024 ** 3)

mapa_ddr = {
    20: "DDR",
    21: "DDR2",
    24: "DDR3",
    26: "DDR4",
    34: "DDR5"
}

tipos_encontrados = set()

for ram in memorias:
    tipo = int(ram.SMBIOSMemoryType or 0)
    
    if tipo == 0:
        tipo = int(ram.MemoryType or 0)  # fallback
    
    if tipo in mapa_ddr:
        tipos_encontrados.add(mapa_ddr[tipo])

if tipos_encontrados:
    tipo_ram = ", ".join(sorted(tipos_encontrados))
else:
    tipo_ram = "Desconhecido"

# Disco
disco = c.Win32_DiskDrive()[0]
capacidade = int(disco.Size)
capacidade_gb = capacidade / (1024 ** 3)

disco_total = f"{capacidade_gb:.2f} GB "

memoria_total = f"{memoria_gb:.0f} GB"