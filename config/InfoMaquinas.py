import wmi

c = wmi.WMI()

# =========================
# 💻 TIPO DE MÁQUINA
# =========================
try:
    chassis = c.Win32_SystemEnclosure()[0].ChassisTypes or []
except:
    chassis = []

notebook_types = {8, 9, 10, 14}

tipo_pc = "Notebook" if any(t in notebook_types for t in chassis) else "Desktop"


# =========================
# 🖥️ SISTEMA
# =========================
system = c.Win32_ComputerSystem()[0]

so = c.Win32_OperatingSystem()[0]
nome_software = so.Caption.replace("Microsoft ", "")


# =========================
# 🧠 PROCESSADOR
# =========================
p = c.Win32_Processor()[0]
nome_processador = " ".join(p.Name.split())


# =========================
# 💾 MEMÓRIA RAM
# =========================
memorias = c.Win32_PhysicalMemory()

memoria_bytes = 0
tipos_encontrados = set()

mapa_ddr = {
    20: "DDR",
    21: "DDR2",
    24: "DDR3",
    26: "DDR4",
    34: "DDR5"
}

for ram in memorias:
    try:
        memoria_bytes += int(ram.Capacity or 0)

        tipo = int(ram.SMBIOSMemoryType or ram.MemoryType or 0)
        if tipo in mapa_ddr:
            tipos_encontrados.add(mapa_ddr[tipo])

    except:
        continue

memoria_gb = memoria_bytes / (1024 ** 3)

tipo_ram = ", ".join(sorted(tipos_encontrados)) if tipos_encontrados else "Desconhecido"
memoria_total = f"{memoria_gb:.0f} GB"


# =========================
# 💽 DISCO (todos os discos)
# =========================
discos = c.Win32_DiskDrive()

capacidade_total = 0

for d in discos:
    try:
        capacidade_total += int(d.Size or 0)
    except:
        continue

capacidade_gb = capacidade_total / (1024 ** 3)

disco_total = f"{capacidade_gb:.2f} GB"