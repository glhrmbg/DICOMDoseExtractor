import pydicom
import datetime
import os


def extrair_tudo_dicom(caminho_arquivo):
    try:
        # Carregar o arquivo DICOM
        ds = pydicom.dcmread(caminho_arquivo)

        # Nome do arquivo de saída baseado no arquivo original
        nome_base = os.path.basename(caminho_arquivo)
        arquivo_saida = f"dicom_completo_{nome_base}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

        with open(arquivo_saida, "w", encoding="utf-8") as f:
            # Cabeçalho do arquivo
            f.write("=" * 80 + "\n")
            f.write("EXTRAÇÃO COMPLETA DE ARQUIVO DICOM\n")
            f.write("=" * 80 + "\n")
            f.write(f"Arquivo: {caminho_arquivo}\n")
            f.write(f"Data da extração: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")

            # SEÇÃO 1: INFORMAÇÕES RESUMIDAS
            f.write("1. RESUMO DAS INFORMAÇÕES PRINCIPAIS\n")
            f.write("-" * 50 + "\n")

            campos_principais = [
                ('PatientName', 'Nome do Paciente'),
                ('PatientID', 'ID do Paciente'),
                ('PatientBirthDate', 'Data de Nascimento'),
                ('PatientSex', 'Sexo'),
                ('Modality', 'Modalidade'),
                ('StudyDescription', 'Descrição do Estudo'),
                ('StudyDate', 'Data do Estudo'),
                ('StudyTime', 'Hora do Estudo'),
                ('SeriesDescription', 'Descrição da Série'),
                ('InstitutionName', 'Instituição'),
                ('ReferringPhysicianName', 'Médico Responsável'),
                ('Manufacturer', 'Fabricante'),
                ('ManufacturerModelName', 'Modelo do Equipamento'),
                ('SoftwareVersions', 'Software')
            ]

            for campo, descricao in campos_principais:
                valor = ds.get(campo, 'N/A')
                f.write(f"{descricao}: {valor}\n")

            f.write("\n" + "=" * 80 + "\n\n")

            # SEÇÃO 2: TODOS OS ELEMENTOS DICOM ORGANIZADOS
            f.write("2. TODOS OS ELEMENTOS DICOM (ORGANIZADOS)\n")
            f.write("-" * 50 + "\n")

            # Agrupar elementos por categoria
            elementos_ordenados = []
            for elem in ds:
                if elem.VR != 'SQ':  # Pular sequências muito complexas por enquanto
                    try:
                        # SEM limitação de tamanho - queremos TUDO!
                        valor_str = str(elem.value)

                        elementos_ordenados.append({
                            'tag': str(elem.tag),
                            'keyword': elem.keyword if elem.keyword else 'SEM_NOME',
                            'vr': elem.VR,
                            'valor': valor_str,
                            'descricao': elem.name if hasattr(elem, 'name') else ''
                        })
                    except:
                        elementos_ordenados.append({
                            'tag': str(elem.tag),
                            'keyword': elem.keyword if elem.keyword else 'SEM_NOME',
                            'vr': elem.VR,
                            'valor': '[ERRO AO CONVERTER]',
                            'descricao': ''
                        })

            # Ordenar por keyword
            elementos_ordenados.sort(key=lambda x: x['keyword'])

            for elem in elementos_ordenados:
                f.write(f"Tag: {elem['tag']} | VR: {elem['vr']} | Keyword: {elem['keyword']}\n")
                f.write(f"Valor: {elem['valor']}\n")
                if elem['descricao']:
                    f.write(f"Descrição: {elem['descricao']}\n")
                f.write("-" * 40 + "\n")

            f.write("\n" + "=" * 80 + "\n\n")

            # SEÇÃO 3: REPRESENTAÇÃO BRUTA COMPLETA
            f.write("3. REPRESENTAÇÃO BRUTA COMPLETA DO DATASET\n")
            f.write("-" * 50 + "\n")
            f.write(str(ds))

            f.write("\n\n" + "=" * 80 + "\n")

            # SEÇÃO 4: SEQUÊNCIAS COMPLEXAS (se existirem)
            f.write("4. SEQUÊNCIAS COMPLEXAS\n")
            f.write("-" * 50 + "\n")

            sequencias_encontradas = False
            for elem in ds:
                if elem.VR == 'SQ':  # Sequências
                    sequencias_encontradas = True
                    f.write(f"\nSequência: {elem.keyword} (Tag: {elem.tag})\n")
                    f.write(f"Número de itens: {len(elem.value) if elem.value else 0}\n")

                    if elem.value:
                        for i, item in enumerate(elem.value):  # TODOS os itens, sem limitação
                            f.write(f"\n  Item {i + 1}:\n")
                            f.write(f"  {str(item)}\n")

                    f.write("-" * 30 + "\n")

            if not sequencias_encontradas:
                f.write("Nenhuma sequência complexa encontrada.\n")

            f.write("\n" + "=" * 80 + "\n")
            f.write("FIM DA EXTRAÇÃO\n")
            f.write("=" * 80 + "\n")

        print(f"✓ Extração completa salva em: {arquivo_saida}")
        print(f"✓ Total de elementos processados: {len(ds)}")
        return arquivo_saida

    except Exception as e:
        print(f"Erro ao processar arquivo DICOM: {e}")
        return None


# Função para usar
def processar_arquivo(caminho):
    """Função principal para processar o arquivo DICOM"""
    if not os.path.exists(caminho):
        print(f"Arquivo não encontrado: {caminho}")
        return

    print(f"Processando arquivo: {caminho}")
    arquivo_gerado = extrair_tudo_dicom(caminho)

    if arquivo_gerado:
        print(f"\nArquivo gerado com sucesso!")
        print(f"Localização: {os.path.abspath(arquivo_gerado)}")

        # Mostrar tamanho do arquivo gerado
        tamanho = os.path.getsize(arquivo_gerado)
        print(f"Tamanho: {tamanho:,} bytes ({tamanho / 1024:.1f} KB)")


# EXEMPLO DE USO:
if __name__ == "__main__":
    # Substitua pelo caminho do seu arquivo
    caminho_arquivo = "seu_arquivo_dicom_aqui"
    processar_arquivo(caminho_arquivo)
