"""
DICOMDoseExtractor.py - Extrator direto de dados de dose de CT de arquivos DICOM SR

Este script lê arquivos DICOM de Structured Report (SR) de dose de radiação
e extrai as informações essenciais diretamente, gerando JSON compatível com
o CTDoseExcel.py existente.
"""

import pydicom
import os
import glob
import json
import argparse
from datetime import datetime
from dataclasses import dataclass, asdict
from typing import List, Optional, Dict, Any


@dataclass
class EssentialInfo:
    """Informações essenciais extraídas do DICOM"""
    patient_id: str = ""
    patient_name: str = ""
    study_id: str = ""
    accession_number: str = ""
    study_date: str = ""
    birth_date: str = ""
    sex: str = ""


@dataclass
class XRaySourceParams:
    """Parâmetros da fonte de raios-X"""
    identification: str = ""
    kvp: str = ""
    max_tube_current: str = ""
    tube_current: str = ""
    exposure_time_per_rotation: Optional[str] = None


@dataclass
class CTDose:
    """Dados de dose CT"""
    mean_ctdivol: str = ""
    phantom_type: str = ""
    dlp: str = ""
    size_specific_dose: Optional[str] = None
    ctdivol_alert_value: Optional[str] = None


@dataclass
class CTAcquisitionParams:
    """Parâmetros de aquisição CT"""
    exposure_time: str = ""
    scanning_length: str = ""
    nominal_single_collimation: str = ""
    nominal_total_collimation: str = ""
    num_xray_sources: str = ""
    pitch_factor: Optional[str] = None


@dataclass
class CTAcquisition:
    """Dados de uma aquisição CT"""
    protocol: str = ""
    target_region: str = ""
    acquisition_type: str = ""
    procedure_context: str = ""
    irradiation_event_uid: str = ""
    comment: str = ""
    acquisition_params: CTAcquisitionParams = None
    xray_source_params: XRaySourceParams = None
    ct_dose: CTDose = None


@dataclass
class IrradiationInfo:
    """Informações de irradiação acumulada"""
    start_time: str = ""
    end_time: str = ""
    total_events: str = ""
    total_dlp: str = ""


@dataclass
class DeviceInfo:
    """Informações do equipamento"""
    observer_name: str = ""
    manufacturer: str = ""
    model_name: str = ""
    serial_number: str = ""
    physical_location: str = ""


@dataclass
class CTScanReport:
    """Relatório completo de dose CT"""
    hospital: str = ""
    report_date: str = ""
    essential: EssentialInfo = None
    device: DeviceInfo = None
    irradiation: IrradiationInfo = None
    acquisitions: List[CTAcquisition] = None

    def __post_init__(self):
        if self.essential is None:
            self.essential = EssentialInfo()
        if self.device is None:
            self.device = DeviceInfo()
        if self.irradiation is None:
            self.irradiation = IrradiationInfo()
        if self.acquisitions is None:
            self.acquisitions = []


class DICOMDoseExtractor:
    """Extrator de dados de dose diretamente de arquivos DICOM SR"""

    def __init__(self):
        # Códigos DICOM para identificação dos campos
        self.concept_codes = {
            # Contexto do dispositivo
            'device_observer_name': '121013',
            'device_observer_manufacturer': '121014',
            'device_observer_model': '121015',
            'device_observer_serial': '121016',
            'device_observer_location': '121017',

            # Dados de irradiação
            'start_irradiation': '113809',
            'end_irradiation': '113810',
            'total_events': '113812',
            'total_dlp': '113813',

            # Aquisição CT
            'ct_acquisition': '113819',
            'acquisition_protocol': '125203',
            'target_region': '123014',
            'acquisition_type': '113820',
            'procedure_context': 'G-C32C',
            'irradiation_event_uid': '113769',
            'comment': '121106',

            # Parâmetros de aquisição
            'acquisition_params': '113822',
            'exposure_time': '113824',
            'scanning_length': '113825',
            'single_collimation': '113826',
            'total_collimation': '113827',
            'num_xray_sources': '113823',
            'pitch_factor': '113828',

            # Parâmetros da fonte de raios-X
            'xray_source_params': '113831',
            'xray_source_id': '113832',
            'kvp': '113733',
            'max_tube_current': '113833',
            'tube_current': '113734',
            'exposure_time_per_rotation': '113834',

            # Dados de dose
            'ct_dose': '113829',
            'mean_ctdivol': '113830',
            'phantom_type': '113835',
            'dlp': '113838',
            'ssde': '113930',
            'ctdivol_alert_value': '113904'
        }

    def format_datetime(self, dt_str: str) -> str:
        """Converte DICOM datetime para formato legível"""
        if not dt_str:
            return ""

        try:
            # DICOM datetime format: YYYYMMDDHHMMSS.FFFFFF
            if '.' in dt_str:
                dt_part = dt_str.split('.')[0]
            else:
                dt_part = dt_str

            if len(dt_part) >= 8:
                year = dt_part[:4]
                month = dt_part[4:6]
                day = dt_part[6:8]

                if len(dt_part) >= 14:
                    hour = dt_part[8:10]
                    minute = dt_part[10:12]
                    second = dt_part[12:14]
                    return f"{year}-{month}-{day} {hour}:{minute}:{second}"
                else:
                    return f"{year}-{month}-{day}"
        except:
            pass

        return dt_str

    def format_date(self, date_str: str) -> str:
        """Converte DICOM date para formato legível"""
        if not date_str or len(date_str) < 8:
            return ""

        try:
            year = date_str[:4]
            month = date_str[4:6]
            day = date_str[6:8]

            # Converte para formato mais legível
            months = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            month_name = months[int(month)]
            return f"{month_name} {int(day)}, {year}"
        except:
            return date_str

    def get_numeric_value_with_unit(self, content_item) -> str:
        """Extrai valor numérico com unidade de uma sequência MeasuredValue"""
        try:
            if hasattr(content_item, 'MeasuredValueSequence') and content_item.MeasuredValueSequence:
                measured_value = content_item.MeasuredValueSequence[0]

                # Valor numérico
                numeric_value = getattr(measured_value, 'NumericValue', '')

                # Unidade
                unit = ''
                if hasattr(measured_value,
                           'MeasurementUnitsCodeSequence') and measured_value.MeasurementUnitsCodeSequence:
                    unit_seq = measured_value.MeasurementUnitsCodeSequence[0]
                    unit = getattr(unit_seq, 'CodeMeaning', '')

                if numeric_value and unit:
                    return f"{numeric_value} {unit}"
                elif numeric_value:
                    return str(numeric_value)

        except Exception as e:
            pass

        return ""

    def get_text_value(self, content_item) -> str:
        """Extrai valor de texto"""
        return getattr(content_item, 'TextValue', '')

    def get_code_meaning(self, content_item) -> str:
        """Extrai Code Meaning de uma sequência de conceito"""
        try:
            if hasattr(content_item, 'ConceptCodeSequence') and content_item.ConceptCodeSequence:
                return getattr(content_item.ConceptCodeSequence[0], 'CodeMeaning', '')
        except:
            pass
        return ""

    def get_datetime_value(self, content_item) -> str:
        """Extrai valor datetime"""
        dt_value = getattr(content_item, 'DateTime', '')
        return self.format_datetime(dt_value)

    def find_content_by_code(self, content_sequence, code_value: str):
        """Encontra item na ContentSequence pelo código do conceito"""
        for item in content_sequence:
            try:
                if hasattr(item, 'ConceptNameCodeSequence') and item.ConceptNameCodeSequence:
                    concept_code = getattr(item.ConceptNameCodeSequence[0], 'CodeValue', '')
                    if concept_code == code_value:
                        return item
            except:
                continue
        return None

    def extract_patient_info(self, ds) -> EssentialInfo:
        """Extrai informações básicas do paciente"""
        essential = EssentialInfo()

        # Patient ID
        essential.patient_id = str(getattr(ds, 'PatientID', ''))

        # Patient Name (limpa formatação DICOM)
        patient_name = str(getattr(ds, 'PatientName', ''))
        if patient_name:
            # Remove os caracteres ^^^^ do final se existirem
            essential.patient_name = patient_name.replace('^', ' ').strip()

        # Study ID
        essential.study_id = str(getattr(ds, 'StudyID', ''))

        # Accession Number
        essential.accession_number = str(getattr(ds, 'AccessionNumber', ''))

        # Study Date
        study_date = str(getattr(ds, 'StudyDate', ''))
        study_time = str(getattr(ds, 'StudyTime', ''))
        if study_date:
            formatted_date = self.format_date(study_date)
            if study_time and len(study_time) >= 6:
                # Adiciona horário se disponível
                hour = study_time[:2]
                minute = study_time[2:4]
                second = study_time[4:6]
                essential.study_date = f"{formatted_date}, {hour}:{minute}:{second}"
            else:
                essential.study_date = formatted_date

        # Birth Date
        birth_date = str(getattr(ds, 'PatientBirthDate', ''))
        if birth_date:
            essential.birth_date = self.format_date(birth_date)

        # Sex
        essential.sex = str(getattr(ds, 'PatientSex', ''))

        return essential

    def extract_device_info(self, content_sequence) -> DeviceInfo:
        """Extrai informações do dispositivo"""
        device = DeviceInfo()

        # Device Observer Name
        item = self.find_content_by_code(content_sequence, self.concept_codes['device_observer_name'])
        if item:
            device.observer_name = self.get_text_value(item)

        # Manufacturer
        item = self.find_content_by_code(content_sequence, self.concept_codes['device_observer_manufacturer'])
        if item:
            device.manufacturer = self.get_text_value(item)

        # Model Name
        item = self.find_content_by_code(content_sequence, self.concept_codes['device_observer_model'])
        if item:
            device.model_name = self.get_text_value(item)

        # Serial Number
        item = self.find_content_by_code(content_sequence, self.concept_codes['device_observer_serial'])
        if item:
            device.serial_number = self.get_text_value(item)

        # Physical Location
        item = self.find_content_by_code(content_sequence, self.concept_codes['device_observer_location'])
        if item:
            device.physical_location = self.get_text_value(item)

        return device

    def extract_irradiation_info(self, content_sequence) -> IrradiationInfo:
        """Extrai informações de irradiação acumulada"""
        irradiation = IrradiationInfo()

        # Start time
        item = self.find_content_by_code(content_sequence, self.concept_codes['start_irradiation'])
        if item:
            irradiation.start_time = self.get_datetime_value(item)

        # End time
        item = self.find_content_by_code(content_sequence, self.concept_codes['end_irradiation'])
        if item:
            irradiation.end_time = self.get_datetime_value(item)

        # Procura por container de dados acumulados
        for item in content_sequence:
            try:
                if (hasattr(item, 'ConceptNameCodeSequence') and
                        item.ConceptNameCodeSequence and
                        getattr(item.ConceptNameCodeSequence[0], 'CodeValue',
                                '') == '113811'):  # CT Accumulated Dose Data

                    if hasattr(item, 'ContentSequence'):
                        # Total events
                        events_item = self.find_content_by_code(item.ContentSequence,
                                                                self.concept_codes['total_events'])
                        if events_item:
                            irradiation.total_events = self.get_numeric_value_with_unit(events_item)

                        # Total DLP
                        dlp_item = self.find_content_by_code(item.ContentSequence, self.concept_codes['total_dlp'])
                        if dlp_item:
                            irradiation.total_dlp = self.get_numeric_value_with_unit(dlp_item)
                    break
            except:
                continue

        return irradiation

    def extract_acquisition_params(self, content_sequence) -> CTAcquisitionParams:
        """Extrai parâmetros de aquisição CT"""
        params = CTAcquisitionParams()

        # Exposure Time
        item = self.find_content_by_code(content_sequence, self.concept_codes['exposure_time'])
        if item:
            params.exposure_time = self.get_numeric_value_with_unit(item)

        # Scanning Length
        item = self.find_content_by_code(content_sequence, self.concept_codes['scanning_length'])
        if item:
            params.scanning_length = self.get_numeric_value_with_unit(item)

        # Single Collimation
        item = self.find_content_by_code(content_sequence, self.concept_codes['single_collimation'])
        if item:
            params.nominal_single_collimation = self.get_numeric_value_with_unit(item)

        # Total Collimation
        item = self.find_content_by_code(content_sequence, self.concept_codes['total_collimation'])
        if item:
            params.nominal_total_collimation = self.get_numeric_value_with_unit(item)

        # Number of X-Ray Sources
        item = self.find_content_by_code(content_sequence, self.concept_codes['num_xray_sources'])
        if item:
            params.num_xray_sources = self.get_numeric_value_with_unit(item)

        # Pitch Factor
        item = self.find_content_by_code(content_sequence, self.concept_codes['pitch_factor'])
        if item:
            params.pitch_factor = self.get_numeric_value_with_unit(item)

        return params

    def extract_xray_source_params(self, content_sequence) -> XRaySourceParams:
        """Extrai parâmetros da fonte de raios-X"""
        xray_params = XRaySourceParams()

        # Source ID
        item = self.find_content_by_code(content_sequence, self.concept_codes['xray_source_id'])
        if item:
            xray_params.identification = self.get_text_value(item)

        # KVP
        item = self.find_content_by_code(content_sequence, self.concept_codes['kvp'])
        if item:
            xray_params.kvp = self.get_numeric_value_with_unit(item)

        # Max Tube Current
        item = self.find_content_by_code(content_sequence, self.concept_codes['max_tube_current'])
        if item:
            xray_params.max_tube_current = self.get_numeric_value_with_unit(item)

        # Tube Current
        item = self.find_content_by_code(content_sequence, self.concept_codes['tube_current'])
        if item:
            xray_params.tube_current = self.get_numeric_value_with_unit(item)

        # Exposure Time per Rotation
        item = self.find_content_by_code(content_sequence, self.concept_codes['exposure_time_per_rotation'])
        if item:
            xray_params.exposure_time_per_rotation = self.get_numeric_value_with_unit(item)

        return xray_params

    def extract_ct_dose(self, content_sequence) -> CTDose:
        """Extrai dados de dose CT"""
        dose = CTDose()

        # Mean CTDIvol
        item = self.find_content_by_code(content_sequence, self.concept_codes['mean_ctdivol'])
        if item:
            dose.mean_ctdivol = self.get_numeric_value_with_unit(item)

        # Phantom Type
        item = self.find_content_by_code(content_sequence, self.concept_codes['phantom_type'])
        if item:
            dose.phantom_type = self.get_code_meaning(item)

        # DLP
        item = self.find_content_by_code(content_sequence, self.concept_codes['dlp'])
        if item:
            dose.dlp = self.get_numeric_value_with_unit(item)

        # Size Specific Dose Estimation (SSDE)
        item = self.find_content_by_code(content_sequence, self.concept_codes['ssde'])
        if item:
            dose.size_specific_dose = self.get_numeric_value_with_unit(item)

        # CTDIvol Alert Value
        item = self.find_content_by_code(content_sequence, self.concept_codes['ctdivol_alert_value'])
        if item:
            dose.ctdivol_alert_value = self.get_numeric_value_with_unit(item)

        return dose

    def extract_ct_acquisitions(self, content_sequence) -> List[CTAcquisition]:
        """Extrai todas as aquisições CT"""
        acquisitions = []

        # Procura por containers de aquisição CT
        for item in content_sequence:
            try:
                if (hasattr(item, 'ConceptNameCodeSequence') and
                        item.ConceptNameCodeSequence and
                        getattr(item.ConceptNameCodeSequence[0], 'CodeValue', '') == self.concept_codes[
                            'ct_acquisition']):

                    acquisition = CTAcquisition()

                    if hasattr(item, 'ContentSequence'):
                        acq_content = item.ContentSequence

                        # Acquisition Protocol
                        protocol_item = self.find_content_by_code(acq_content,
                                                                  self.concept_codes['acquisition_protocol'])
                        if protocol_item:
                            acquisition.protocol = self.get_text_value(protocol_item)

                        # Target Region
                        target_item = self.find_content_by_code(acq_content, self.concept_codes['target_region'])
                        if target_item:
                            acquisition.target_region = self.get_code_meaning(target_item)

                        # Acquisition Type
                        type_item = self.find_content_by_code(acq_content, self.concept_codes['acquisition_type'])
                        if type_item:
                            acquisition.acquisition_type = self.get_code_meaning(type_item)

                        # Procedure Context
                        context_item = self.find_content_by_code(acq_content, self.concept_codes['procedure_context'])
                        if context_item:
                            acquisition.procedure_context = self.get_code_meaning(context_item)

                        # Irradiation Event UID
                        uid_item = self.find_content_by_code(acq_content, self.concept_codes['irradiation_event_uid'])
                        if uid_item:
                            acquisition.irradiation_event_uid = str(getattr(uid_item, 'UID', ''))

                        # Comment
                        comment_item = self.find_content_by_code(acq_content, self.concept_codes['comment'])
                        if comment_item:
                            acquisition.comment = self.get_text_value(comment_item)

                        # Procura por sub-containers
                        for sub_item in acq_content:
                            try:
                                if (hasattr(sub_item, 'ConceptNameCodeSequence') and
                                        sub_item.ConceptNameCodeSequence):

                                    code = getattr(sub_item.ConceptNameCodeSequence[0], 'CodeValue', '')

                                    # Acquisition Parameters
                                    if code == self.concept_codes['acquisition_params'] and hasattr(sub_item,
                                                                                                    'ContentSequence'):
                                        acquisition.acquisition_params = self.extract_acquisition_params(
                                            sub_item.ContentSequence)

                                        # Dentro dos params, procura por X-Ray Source Parameters
                                        for param_item in sub_item.ContentSequence:
                                            if (hasattr(param_item, 'ConceptNameCodeSequence') and
                                                    param_item.ConceptNameCodeSequence and
                                                    getattr(param_item.ConceptNameCodeSequence[0], 'CodeValue', '') ==
                                                    self.concept_codes['xray_source_params'] and
                                                    hasattr(param_item, 'ContentSequence')):
                                                acquisition.xray_source_params = self.extract_xray_source_params(
                                                    param_item.ContentSequence)
                                                break

                                    # CT Dose
                                    elif code == self.concept_codes['ct_dose'] and hasattr(sub_item, 'ContentSequence'):
                                        acquisition.ct_dose = self.extract_ct_dose(sub_item.ContentSequence)

                            except:
                                continue

                    acquisitions.append(acquisition)

            except:
                continue

        return acquisitions

    def extract_from_dicom(self, dicom_path: str, debug_mode: bool = False) -> CTScanReport:
        """Extrai informações de um arquivo DICOM SR"""

        if debug_mode:
            print(f"\n{'=' * 80}")
            print(f"PROCESSANDO DICOM: {os.path.basename(dicom_path)}")
            print(f"{'=' * 80}")

        try:
            # Lê o arquivo DICOM
            ds = pydicom.dcmread(dicom_path)

            if debug_mode:
                print(f"SOP Class: {getattr(ds, 'SOPClassUID', 'Unknown')}")
                print(f"Modality: {getattr(ds, 'Modality', 'Unknown')}")

            # Verifica se é um Structured Report de dose
            if (not hasattr(ds, 'Modality') or ds.Modality != 'SR' or
                    not hasattr(ds, 'ContentSequence')):
                if debug_mode:
                    print("❌ Arquivo não é um DICOM SR válido")
                return None

            report = CTScanReport()

            # Extrai informações básicas
            report.essential = self.extract_patient_info(ds)

            # Hospital e data do relatório
            report.hospital = str(getattr(ds, 'InstitutionName', ''))
            content_date = str(getattr(ds, 'ContentDate', ''))
            content_time = str(getattr(ds, 'ContentTime', ''))
            if content_date:
                formatted_date = self.format_date(content_date)
                if content_time and len(content_time) >= 6:
                    hour = content_time[:2]
                    minute = content_time[2:4]
                    second = content_time[4:6]
                    report.report_date = f"{formatted_date}, {hour}:{minute}:{second}"
                else:
                    report.report_date = formatted_date

            # Extrai dados do Content Sequence principal
            main_content = ds.ContentSequence

            # Device info
            report.device = self.extract_device_info(main_content)

            # Irradiation info
            report.irradiation = self.extract_irradiation_info(main_content)

            # CT Acquisitions
            report.acquisitions = self.extract_ct_acquisitions(main_content)

            if debug_mode:
                print(f"✓ Dados extraídos:")
                print(f"  Patient ID: {report.essential.patient_id}")
                print(f"  Patient Name: {report.essential.patient_name}")
                print(f"  Study ID: {report.essential.study_id}")
                print(f"  Hospital: {report.hospital}")
                print(f"  Total DLP: {report.irradiation.total_dlp}")
                print(f"  Acquisições encontradas: {len(report.acquisitions)}")
                for i, acq in enumerate(report.acquisitions, 1):
                    print(f"    {i}. {acq.protocol} - CTDIvol: {acq.ct_dose.mean_ctdivol if acq.ct_dose else 'N/A'}")

            return report

        except Exception as e:
            if debug_mode:
                print(f"❌ Erro ao processar DICOM: {str(e)}")
            return None


def process_dicom_folder(folder_path: str = "ct_dicoms", json_folder: str = "ct_reports_json",
                         debug_mode: bool = False) -> List[Dict]:
    """Processa todos os arquivos DICOM em uma pasta"""

    # Cria a pasta se não existir
    if not os.path.exists(folder_path):
        try:
            os.makedirs(folder_path)
            print(f"✓ Pasta '{folder_path}' criada com sucesso!")
        except Exception as e:
            print(f"✗ Erro ao criar pasta '{folder_path}': {str(e)}")
            return []

    # Busca arquivos DICOM (sem extensão ou .dcm)
    dicom_files = []

    # Arquivos sem extensão
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path) and '.' not in file:
            dicom_files.append(file_path)

    # Arquivos .dcm
    dicom_pattern = os.path.join(folder_path, "*.dcm")
    dicom_files.extend(glob.glob(dicom_pattern))

    if not dicom_files:
        print(f"ℹ️ Nenhum arquivo DICOM encontrado na pasta '{folder_path}'.")
        return []

    print(f"🔍 Encontrados {len(dicom_files)} arquivos DICOM para processar.")

    extractor = DICOMDoseExtractor()
    reports = []

    for dicom_file in dicom_files:
        try:
            if debug_mode:
                print(f"\n{'=' * 80}")
                print(f"PROCESSANDO: {os.path.basename(dicom_file)}")
                print(f"{'=' * 80}")

            report = extractor.extract_from_dicom(dicom_file, debug_mode=debug_mode)

            if report:
                report_dict = asdict(report)
                reports.append(report_dict)

                # Salva o relatório individual usando o Patient ID
                patient_id = report.essential.patient_id
                if patient_id:
                    output_file = f"ct_report_dicom_{patient_id}.json"
                    save_to_json([report_dict], output_file, json_folder)
                    print(
                        f"✓ Processado e salvo: {os.path.basename(dicom_file)} → {os.path.join(json_folder, output_file)}")
                else:
                    print(f"✓ Processado: {os.path.basename(dicom_file)} (sem Patient ID)")
            else:
                print(f"⚠️ Não foi possível extrair dados de: {os.path.basename(dicom_file)}")

        except Exception as e:
            print(f"✗ Erro ao processar {os.path.basename(dicom_file)}: {str(e)}")

    return reports


def save_to_json(reports: List[Dict], output_file: str, json_folder: str = "ct_reports_json"):
    """Salva os relatórios em um arquivo JSON"""
    # Cria a pasta JSON se não existir
    if not os.path.exists(json_folder):
        try:
            os.makedirs(json_folder)
            print(f"✓ Pasta '{json_folder}' criada com sucesso!")
        except Exception as e:
            print(f"✗ Erro ao criar pasta '{json_folder}': {str(e)}")
            return False

    try:
        output_path = os.path.join(json_folder, output_file)

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(reports, f, indent=2, ensure_ascii=False)

        return True

    except Exception as e:
        print(f"✗ Erro ao salvar JSON {output_file}: {str(e)}")
        return False


def main():
    """Função principal"""
    parser = argparse.ArgumentParser(description='DICOMDoseExtractor - Extrai dados de dose de arquivos DICOM SR')
    parser.add_argument('--folder', '-f', default='ct_dicoms',
                        help='Pasta contendo os arquivos DICOM (padrão: ct_dicoms)')
    parser.add_argument('--output-folder', '-o', default='ct_reports_json',
                        help='Pasta para salvar os JSONs (padrão: ct_reports_json)')
    parser.add_argument('--debug', '-d', action='store_true',
                        help='Ativa o modo debug com informações detalhadas')
    parser.add_argument('--single', '-s', type=str,
                        help='Processa um único arquivo DICOM')

    args = parser.parse_args()

    print("=" * 80)
    print("🏥 DICOM DOSE EXTRACTOR - Extrator Direto de Arquivos DICOM")
    print("=" * 80)
    print(f"📂 Pasta de entrada: {args.folder}")
    print(f"📁 Pasta de saída: {args.output_folder}")
    print(f"🔍 Modo debug: {'Ativado' if args.debug else 'Desativado'}")
    print("=" * 80)

    if args.single:
        # Processa um único arquivo
        if not os.path.exists(args.single):
            print(f"❌ Arquivo não encontrado: {args.single}")
            return

        extractor = DICOMDoseExtractor()
        print(f"🔍 Processando arquivo único: {os.path.basename(args.single)}")

        report = extractor.extract_from_dicom(args.single, debug_mode=args.debug)

        if report:
            report_dict = asdict(report)
            patient_id = report.essential.patient_id or "unknown"
            output_file = f"ct_report_dicom_{patient_id}.json"

            if save_to_json([report_dict], output_file, args.output_folder):
                print(f"✅ Relatório salvo em: {os.path.join(args.output_folder, output_file)}")
            else:
                print("❌ Erro ao salvar o relatório")
        else:
            print("❌ Não foi possível extrair dados do arquivo DICOM")
    else:
        # Processa pasta inteira
        reports = process_dicom_folder(args.folder, args.output_folder, args.debug)

        if reports:
            # Salva relatório consolidado
            consolidated_file = f"all_ct_reports_dicom_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            if save_to_json(reports, consolidated_file, args.output_folder):
                print(f"\n✅ Relatório consolidado salvo em: {os.path.join(args.output_folder, consolidated_file)}")

            print(f"\n📊 RESUMO:")
            print(f"   Total de relatórios processados: {len(reports)}")
            print(f"   Arquivos salvos na pasta: {args.output_folder}")

            # Mostra estatísticas básicas
            total_dlp_values = []
            hospitals = set()

            for report in reports:
                if report.get('irradiation', {}).get('total_dlp'):
                    dlp_str = report['irradiation']['total_dlp']
                    try:
                        # Extrai apenas o número da string "valor unidade"
                        dlp_value = float(dlp_str.split()[0])
                        total_dlp_values.append(dlp_value)
                    except:
                        pass

                hospital = report.get('hospital', '')
                if hospital:
                    hospitals.add(hospital)

            if total_dlp_values:
                print(
                    f"   DLP Total - Min: {min(total_dlp_values):.2f}, Max: {max(total_dlp_values):.2f}, Média: {sum(total_dlp_values) / len(total_dlp_values):.2f}")

            if hospitals:
                print(f"   Hospitais: {', '.join(list(hospitals)[:3])}{'...' if len(hospitals) > 3 else ''}")
        else:
            print("\n⚠️ Nenhum relatório foi processado com sucesso.")

    print("\n🎯 Processamento concluído!")


if __name__ == "__main__":
    main()