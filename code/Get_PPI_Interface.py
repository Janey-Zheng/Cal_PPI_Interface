from pymol import cmd
import pymol
from interfaceResidues import interfaceResidues
import os
import openpyxl

# 启动 PyMOL（不显示 GUI）
pymol.finish_launching(['pymol', '-cq'])

pdb_dir = r"F:/Data_0822/database/structures/database_pdb"
output_excel = r"F:/Data_0822/Interface/Interface_Results.xlsx"
cutoff = 1.0  # ΔSASA 阈值

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Interface_Summary"
ws.append(["PDB_ID", "Chain1", "Chain2", "Interface1", "Interface2"])

for filename in os.listdir(pdb_dir):
    # if filename.endswith(".pdb"):
    pdb_id = os.path.splitext(filename)[0]
    pdb_path = os.path.join(pdb_dir, filename)
    print(f"Processing {pdb_id}...")
    
    # 重置 PyMOL 环境
    cmd.reinitialize()
    
    object_name = f"obj_{pdb_id}"
    cmd.load(pdb_path, object_name)
    chains = cmd.get_chains(object_name)
    chain1 = chains[0]
    chain2 = chains[1]
    c1 = f"c. {chain1}"
    c2 = f"c. {chain2}"
    
    # 进行界面残基的识别
    res_list = interfaceResidues(object_name, c1, c2, cutoff)
    interface1 = []
    interface2 = []
    for model, resi, delta in res_list:
        if chain1 in model:
            # interface1.append(int(resi))
            interface1.append(resi)
        elif chain2 in model:
            # interface2.append(int(resi))
            interface2.append(resi)
    # interface1 = sorted(set(interface1))
    interface1 = sorted(set(interface1), key=lambda x: (int(''.join([c for c in x if c.isdigit()]) or 0), x))
    # interface2 = sorted(set(interface2))
    interface2 = sorted(set(interface2), key=lambda x: (int(''.join([c for c in x if c.isdigit()]) or 0), x))
    ws.append([pdb_id, chain1, chain2, str(interface1), str(interface2)])
wb.save(output_excel)

# cmd.load("F:/Data_0822/Interface/str_try/1a1b.pdb", "myComplex")
# chains = cmd.get_chains("myComplex")
# chainA = f"c. {chains[0]}"
# chainB = f"c. {chains[1]}"
# # 分析
# res = interfaceResidues("myComplex", chainA, chainB, 1.0, "intf")
# print(res)
