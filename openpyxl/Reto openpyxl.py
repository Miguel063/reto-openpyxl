import openpyxl
from openpyxl import Workbook
import os

def abr_o_cre(excel):
    if os.path.exists(excel):
        wb = openpyxl.load_workbook(excel)
    else:
        wb = Workbook()
        hoja = wb.active
        hoja.title = "Gastos"
        hoja.append(["Fecha", "Descripción", "Monto (Q)"])
    return wb

def ad_gas(fecha, descripcion, monto, hoja):
    hoja.append([fecha, descripcion, monto])

def resu(gas):
    to_ga = len(gas)
    gas_caro = max(gas, key=lambda x: x[2])
    gas_bar = min(gas, key=lambda x: x[2])
    to_mon = sum(ga[2] for ga in gas)

    print("\nResumen de gastos:")
    print(f"Total de gastos: {to_ga}")
    print(f"Gasto más caro: {gas_caro[1]} el {gas_caro[0]} con un monto de Q{gas_caro[2]:.2f}")
    print(f"Gasto más barato: {gas_bar[1]} el {gas_bar[0]} con un monto de Q{gas_bar[2]:.2f}")
    print(f"Monto total de gastos recien ingresos es: Q{to_mon:.2f}")

def ab_ar(excel):
    os.startfile(excel)

def main():
    excel = "Informe_gastos.xlsx"
    wb = abr_o_cre(excel)
    hoja = wb.active

    gas = []

    while True:
        print("\nIngresa un nuevo gasto:")
        fecha = input("Fecha del gasto (YYYY-MM-DD): ")
        desc = input("Descripción del gasto: ")
        while True:
            try:
                monto = float(input("Monto del gasto (Q): "))
                break
            except ValueError:
                print("Por favor, ingresa un monto válido.")
        
        gas.append([fecha, desc, monto])

        ad_gas(fecha, desc, monto, hoja)

        continuar = input("¿Deseas agregar otro gasto? (si/no): ").lower()
        if continuar != 'si':
            break

    resu(gas)

    wb.save(excel)
    print(f"\nLos datos de los gastos han sido guardados en {excel}")
#puede probar a borrar el excel para crear otro desde 0
    ab_ar(excel)
    
main()
