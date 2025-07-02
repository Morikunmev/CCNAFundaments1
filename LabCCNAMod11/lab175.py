import pandas as pd
import ipaddress

# Crear la estructura base de la tabla
def crear_tabla_red():
    """
    Crea una tabla de configuración de red con dispositivos, interfaces y configuración IP
    """
    
    # Definir los datos base
    dispositivos = [
        # Routers
        {'Dispositivo': 'R1', 'Interfaz': 'G0/0', 'Tipo': 'Router'},
        {'Dispositivo': 'R1', 'Interfaz': 'G0/1', 'Tipo': 'Router'},
        {'Dispositivo': 'R1', 'Interfaz': 'S0/0/0', 'Tipo': 'Router'},
        {'Dispositivo': 'R2', 'Interfaz': 'G0/0', 'Tipo': 'Router'},
        {'Dispositivo': 'R2', 'Interfaz': 'G0/1', 'Tipo': 'Router'},
        {'Dispositivo': 'R2', 'Interfaz': 'S0/0/0', 'Tipo': 'Router'},
        
        # Switches
        {'Dispositivo': 'S1', 'Interfaz': 'VLAN 1', 'Tipo': 'Switch'},
        {'Dispositivo': 'S2', 'Interfaz': 'VLAN 1', 'Tipo': 'Switch'},
        {'Dispositivo': 'S3', 'Interfaz': 'VLAN 1', 'Tipo': 'Switch'},
        {'Dispositivo': 'S4', 'Interfaz': 'VLAN 1', 'Tipo': 'Switch'},
        
        # PCs
        {'Dispositivo': 'PC1', 'Interfaz': 'NIC', 'Tipo': 'PC'},
        {'Dispositivo': 'PC2', 'Interfaz': 'NIC', 'Tipo': 'PC'},
        {'Dispositivo': 'PC3', 'Interfaz': 'NIC', 'Tipo': 'PC'},
        {'Dispositivo': 'PC4', 'Interfaz': 'NIC', 'Tipo': 'PC'},
    ]
    
    # Crear DataFrame
    df = pd.DataFrame(dispositivos)
    
    # Agregar columnas vacías para configuración
    df['Dirección IP'] = ''
    df['Máscara de subred'] = ''
    df['Gateway predeterminado'] = ''
    
    return df

def asignar_ips_automaticamente(df, red_base="192.168.1.0/24"):
    """
    Asigna direcciones IP automáticamente basado en una red base
    """
    red = ipaddress.IPv4Network(red_base, strict=False)
    hosts = list(red.hosts())
    
    # Diccionario para asignar IPs por tipo de dispositivo
    ip_counter = 0
    
    for index, row in df.iterrows():
        dispositivo = row['Dispositivo']
        tipo = row['Tipo']
        
        if tipo == 'Router':
            # Routers usan las primeras IPs
            if 'G0/0' in row['Interfaz']:
                df.at[index, 'Dirección IP'] = str(hosts[ip_counter])
                df.at[index, 'Máscara de subred'] = str(red.netmask)
                ip_counter += 1
            elif 'G0/1' in row['Interfaz']:
                df.at[index, 'Dirección IP'] = str(hosts[ip_counter])
                df.at[index, 'Máscara de subred'] = str(red.netmask)
                ip_counter += 1
            elif 'S0/0/0' in row['Interfaz']:
                # Interfaces seriales pueden usar otra subred
                df.at[index, 'Dirección IP'] = f"10.0.0.{ip_counter}"
                df.at[index, 'Máscara de subred'] = "255.255.255.252"
                ip_counter += 1
                
        elif tipo == 'Switch':
            # Switches para administración
            df.at[index, 'Dirección IP'] = str(hosts[ip_counter + 10])
            df.at[index, 'Máscara de subred'] = str(red.netmask)
            df.at[index, 'Gateway predeterminado'] = str(hosts[0])  # Primer router
            
        elif tipo == 'PC':
            # PCs usan IPs del final del rango
            df.at[index, 'Dirección IP'] = str(hosts[ip_counter + 20])
            df.at[index, 'Máscara de subred'] = str(red.netmask)
            df.at[index, 'Gateway predeterminado'] = str(hosts[0])  # Primer router
            ip_counter += 1
    
    return df

def exportar_a_excel(df, nombre_archivo="configuracion_red.xlsx"):
    """
    Exporta la tabla a un archivo Excel con formato
    """
    with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Configuración de Red', index=False)
        
        # Obtener el workbook y worksheet para formato
        workbook = writer.book
        worksheet = writer.sheets['Configuración de Red']
        
        # Ajustar ancho de columnas
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

def mostrar_tabla(df):
    """
    Muestra la tabla en formato tabular
    """
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 20)
    print("=== CONFIGURACIÓN DE RED ===")
    print(df.to_string(index=False))

def agregar_dispositivo(df, dispositivo, interfaz, tipo):
    """
    Agrega un nuevo dispositivo a la tabla
    """
    nuevo_dispositivo = {
        'Dispositivo': dispositivo,
        'Interfaz': interfaz,
        'Tipo': tipo,
        'Dirección IP': '',
        'Máscara de subred': '',
        'Gateway predeterminado': ''
    }
    
    return pd.concat([df, pd.DataFrame([nuevo_dispositivo])], ignore_index=True)

# Ejemplo de uso
if __name__ == "__main__":
    # Crear tabla base
    tabla = crear_tabla_red()
    
    print("1. Tabla inicial (vacía):")
    mostrar_tabla(tabla)
    
    print("\n" + "="*80 + "\n")
    
    # Asignar IPs automáticamente
    tabla_con_ips = asignar_ips_automaticamente(tabla.copy())
    
    print("2. Tabla con IPs asignadas automáticamente:")
    mostrar_tabla(tabla_con_ips)
    
    print("\n" + "="*80 + "\n")
    
    # Agregar un dispositivo nuevo
    tabla_expandida = agregar_dispositivo(tabla_con_ips, 'R3', 'G0/0', 'Router')
    
    print("3. Tabla con dispositivo agregado:")
    mostrar_tabla(tabla_expandida)
    
    # Exportar a Excel
    try:
        exportar_a_excel(tabla_con_ips)
        print(f"\n✅ Archivo 'configuracion_red.xlsx' creado exitosamente!")
    except Exception as e:
        print(f"\n❌ Error al crear archivo Excel: {e}")
    
    print("\n=== FUNCIONES DISPONIBLES ===")
    print("- crear_tabla_red(): Crea tabla vacía")
    print("- asignar_ips_automaticamente(df, red): Asigna IPs automáticamente") 
    print("- agregar_dispositivo(df, device, interface, type): Agrega dispositivo")
    print("- exportar_a_excel(df, filename): Exporta a Excel")
    print("- mostrar_tabla(df): Muestra tabla formateada")