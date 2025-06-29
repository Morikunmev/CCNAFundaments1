import pandas as pd
import os
from datetime import datetime

class NetworkTablesGenerator:
    def __init__(self):
        self.address_data = [
            {"Dispositivo": "R1", "Interfaz": "G0/0", "Dirección IP": "192.168.10.1", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": "N/A"},
            {"Dispositivo": "R1", "Interfaz": "G0/1", "Dirección IP": "192.168.11.1", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": "N/A"},
            {"Dispositivo": "S1", "Interfaz": "VLAN 1", "Dirección IP": "192.168.10.2", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": ""},
            {"Dispositivo": "S2", "Interfaz": "VLAN 1", "Dirección IP": "192.168.11.2", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": ""},
            {"Dispositivo": "PC1", "Interfaz": "NIC", "Dirección IP": "192.168.10.10", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": ""},
            {"Dispositivo": "PC2", "Interfaz": "NIC", "Dirección IP": "192.168.10.11", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": ""},
            {"Dispositivo": "PC3", "Interfaz": "NIC", "Dirección IP": "192.168.11.10", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": ""},
            {"Dispositivo": "PC4", "Interfaz": "NIC", "Dirección IP": "192.168.11.11", "Máscara de subred": "255.255.255.0", "Puerta de enlace predeterminada": ""}
        ]
        
        self.test_data = [
            {"Prueba": "PC1 a PC2", "¿Se realizó correctamente?": "No", "Problemas": "Dirección IP en la PC1", "Solución": "Cambiar la dirección IP de la PC1", "Verificado": ""},
            {"Prueba": "PC1 a S1", "¿Se realizó correctamente?": "", "Problemas": "", "Solución": "", "Verificado": ""},
            {"Prueba": "PC1 a R1", "¿Se realizó correctamente?": "", "Problemas": "", "Solución": "", "Verificado": ""},
            {"Prueba": "", "¿Se realizó correctamente?": "", "Problemas": "", "Solución": "", "Verificado": ""},
            {"Prueba": "", "¿Se realizó correctamente?": "", "Problemas": "", "Solución": "", "Verificado": ""}
        ]

    def create_dataframes(self):
        """Crea los DataFrames de pandas para ambas tablas"""
        self.df_address = pd.DataFrame(self.address_data)
        self.df_test = pd.DataFrame(self.test_data)
        return self.df_address, self.df_test

    def display_tables(self):
        """Muestra las tablas en consola de forma formateada"""
        print("=" * 80)
        print("TABLA DE ASIGNACIÓN DE DIRECCIONES")
        print("=" * 80)
        print(self.df_address.to_string(index=False))
        
        print("\n" + "=" * 80)
        print("TABLA DE PRUEBAS DE CONECTIVIDAD")
        print("=" * 80)
        print(self.df_test.to_string(index=False))

    def export_to_excel(self, filename=None):
        """Exporta ambas tablas a un archivo Excel con múltiples hojas"""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"tablas_red_{timestamp}.xlsx"
        
        try:
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                # Escribir las tablas en hojas separadas
                self.df_address.to_excel(writer, sheet_name='Asignación_Direcciones', index=False)
                self.df_test.to_excel(writer, sheet_name='Pruebas_Conectividad', index=False)
                
                # Obtener el workbook y worksheets para formateo
                workbook = writer.book
                worksheet1 = writer.sheets['Asignación_Direcciones']
                worksheet2 = writer.sheets['Pruebas_Conectividad']
                
                # Crear formatos
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                cell_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'border': 1
                })
                
                # Aplicar formato a la primera hoja
                for col_num, value in enumerate(self.df_address.columns.values):
                    worksheet1.write(0, col_num, value, header_format)
                    # Ajustar ancho de columnas
                    column_width = max(len(value), 15)
                    worksheet1.set_column(col_num, col_num, column_width)
                
                # Aplicar formato a la segunda hoja
                for col_num, value in enumerate(self.df_test.columns.values):
                    worksheet2.write(0, col_num, value, header_format)
                    # Ajustar ancho de columnas (más ancho para problemas y solución)
                    if col_num in [2, 3]:  # Columnas de problemas y solución
                        column_width = 30
                    else:
                        column_width = max(len(value), 15)
                    worksheet2.set_column(col_num, col_num, column_width)
            
            print(f"\n✅ Archivo Excel creado exitosamente: {filename}")
            return filename
            
        except Exception as e:
            print(f"❌ Error al crear el archivo Excel: {e}")
            return None

    def export_to_csv(self, prefix="tabla_red"):
        """Exporta las tablas a archivos CSV separados"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Exportar tabla de direcciones
            address_filename = f"{prefix}_direcciones_{timestamp}.csv"
            self.df_address.to_csv(address_filename, index=False, encoding='utf-8-sig')
            
            # Exportar tabla de pruebas
            test_filename = f"{prefix}_pruebas_{timestamp}.csv"
            self.df_test.to_csv(test_filename, index=False, encoding='utf-8-sig')
            
            print(f"✅ Archivos CSV creados:")
            print(f"   - {address_filename}")
            print(f"   - {test_filename}")
            
            return address_filename, test_filename
            
        except Exception as e:
            print(f"❌ Error al crear los archivos CSV: {e}")
            return None, None

    def add_device(self, dispositivo, interfaz, ip, mascara, gateway=""):
        """Agrega un nuevo dispositivo a la tabla de direcciones"""
        new_device = {
            "Dispositivo": dispositivo,
            "Interfaz": interfaz,
            "Dirección IP": ip,
            "Máscara de subred": mascara,
            "Puerta de enlace predeterminada": gateway
        }
        self.address_data.append(new_device)
        self.df_address = pd.DataFrame(self.address_data)
        print(f"✅ Dispositivo {dispositivo} agregado exitosamente")

    def add_test(self, prueba, realizado="", problemas="", solucion="", verificado=""):
        """Agrega una nueva prueba a la tabla de pruebas"""
        new_test = {
            "Prueba": prueba,
            "¿Se realizó correctamente?": realizado,
            "Problemas": problemas,
            "Solución": solucion,
            "Verificado": verificado
        }
        self.test_data.append(new_test)
        self.df_test = pd.DataFrame(self.test_data)
        print(f"✅ Prueba '{prueba}' agregada exitosamente")

    def update_gateway(self, dispositivo, gateway):
        """Actualiza el gateway de un dispositivo específico"""
        for item in self.address_data:
            if item["Dispositivo"] == dispositivo:
                item["Puerta de enlace predeterminada"] = gateway
                print(f"✅ Gateway actualizado para {dispositivo}: {gateway}")
                break
        else:
            print(f"❌ Dispositivo {dispositivo} no encontrado")
        
        self.df_address = pd.DataFrame(self.address_data)

def main():
    """Función principal - Ejemplo de uso"""
    print("🌐 GENERADOR DE TABLAS DE RED")
    print("=" * 50)
    
    # Crear instancia del generador
    generator = NetworkTablesGenerator()
    
    # Crear DataFrames
    df_address, df_test = generator.create_dataframes()
    
    # Mostrar las tablas
    generator.display_tables()
    
    # Menú interactivo
    while True:
        print("\n" + "="*50)
        print("OPCIONES:")
        print("1. Agregar dispositivo")
        print("2. Agregar prueba")
        print("3. Actualizar gateway")
        print("4. Exportar a Excel")
        print("5. Exportar a CSV")
        print("6. Mostrar tablas")
        print("7. Salir")
        
        opcion = input("\nSelecciona una opción (1-7): ").strip()
        
        if opcion == "1":
            dispositivo = input("Dispositivo: ")
            interfaz = input("Interfaz: ")
            ip = input("Dirección IP: ")
            mascara = input("Máscara de subred: ")
            gateway = input("Gateway (opcional): ")
            generator.add_device(dispositivo, interfaz, ip, mascara, gateway)
            
        elif opcion == "2":
            prueba = input("Descripción de la prueba: ")
            realizado = input("¿Se realizó correctamente? (Sí/No): ")
            problemas = input("Problemas (opcional): ")
            solucion = input("Solución (opcional): ")
            verificado = input("Verificado (Sí/No/Pendiente): ")
            generator.add_test(prueba, realizado, problemas, solucion, verificado)
            
        elif opcion == "3":
            dispositivo = input("Dispositivo a actualizar: ")
            gateway = input("Nuevo gateway: ")
            generator.update_gateway(dispositivo, gateway)
            
        elif opcion == "4":
            filename = input("Nombre del archivo (opcional, presiona Enter para auto-generar): ").strip()
            if not filename:
                filename = None
            generator.export_to_excel(filename)
            
        elif opcion == "5":
            prefix = input("Prefijo para archivos (opcional, presiona Enter para 'tabla_red'): ").strip()
            if not prefix:
                prefix = "tabla_red"
            generator.export_to_csv(prefix)
            
        elif opcion == "6":
            generator.display_tables()
            
        elif opcion == "7":
            print("¡Hasta luego! 👋")
            break
            
        else:
            print("❌ Opción no válida. Intenta de nuevo.")

if __name__ == "__main__":
    # Verificar si pandas está instalado
    try:
        import pandas as pd
        main()
    except ImportError:
        print("❌ Error: pandas no está instalado.")
        print("📦 Instala pandas con: pip install pandas")
        print("📦 Para Excel también instala: pip install xlsxwriter")