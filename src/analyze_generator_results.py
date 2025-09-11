#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PyPSA-HD 발전량 분석 스크립트

최적화 결과 파일에서 발전기별 출력을 분석하고, 설비별/지역별 발전량을 집계하여 보여줍니다.
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import glob
from datetime import datetime
import json
import re


def find_latest_results():
    """최신 결과 디렉토리를 찾습니다."""
    results_dir = 'results'
    if not os.path.exists(results_dir):
        print(f"결과 디렉토리 '{results_dir}'를 찾을 수 없습니다.")
        return None

    # 결과 디렉토리에서 CSV 파일 찾기
    csv_files = glob.glob(os.path.join(results_dir, '*', '*.csv'))
    
    if not csv_files:
        print("결과 CSV 파일을 찾을 수 없습니다.")
        return None
    
    # 파일 경로에서 최신 결과 디렉토리 찾기
    latest_dir = None
    latest_time = None
    
    for file_path in csv_files:
        # 디렉토리 이름에서 타임스탬프 추출
        dir_name = os.path.basename(os.path.dirname(file_path))
        match = re.search(r'(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})', dir_name)
        
        if match:
            timestamp_str = match.group(1)
            try:
                timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d_%H-%M-%S')
                if latest_time is None or timestamp > latest_time:
                    latest_time = timestamp
                    latest_dir = os.path.dirname(file_path)
            except ValueError:
                continue
    
    if latest_dir:
        print(f"최신 결과 디렉토리를 찾았습니다: {latest_dir}")
        return latest_dir
    else:
        print("결과 디렉토리를 찾을 수 없습니다.")
        return None


def load_generator_output(results_dir):
    """발전기 출력 데이터를 로드합니다."""
    generator_output_file = os.path.join(results_dir, 'generators_output.csv')
    
    if not os.path.exists(generator_output_file):
        print(f"발전기 출력 파일을 찾을 수 없습니다: {generator_output_file}")
        return None
    
    try:
        gen_output = pd.read_csv(generator_output_file, index_col=0)
        print(f"발전기 출력 데이터를 로드했습니다. 크기: {gen_output.shape}")
        
        # 시간 인덱스를 날짜/시간 형식으로 변환
        if gen_output.index.dtype == 'object':
            try:
                gen_output.index = pd.to_datetime(gen_output.index)
                print(f"시간 범위: {gen_output.index[0]} ~ {gen_output.index[-1]}")
            except:
                print("시간 인덱스를 날짜/시간 형식으로 변환할 수 없습니다.")
        
        return gen_output
    except Exception as e:
        print(f"발전기 출력 데이터 로드 중 오류 발생: {str(e)}")
        return None


def load_generator_info(results_dir):
    """발전기 정보를 로드합니다."""
    try:
        # 발전기 정보 파일 찾기
        generator_info_file = os.path.join(results_dir, 'generators_info.csv')
        network_file = os.path.join(results_dir, 'network.csv')
        
        if os.path.exists(generator_info_file):
            gen_info = pd.read_csv(generator_info_file)
            print(f"발전기 정보 데이터를 로드했습니다. 크기: {gen_info.shape}")
            return gen_info
        elif os.path.exists(network_file):
            network = pd.read_csv(network_file)
            if 'Generator' in network['component'].values:
                gen_info = network[network['component'] == 'Generator']
                print(f"네트워크 파일에서 발전기 정보를 로드했습니다. 크기: {gen_info.shape}")
                return gen_info
        
        print("발전기 정보 파일이나 네트워크 파일을 찾을 수 없습니다.")
        return None
    except Exception as e:
        print(f"발전기 정보 로드 중 오류 발생: {str(e)}")
        return None


def infer_carrier_from_name(name):
    """발전기 이름에서 에너지원(carrier) 유형을 추론합니다."""
    name_lower = name.lower()
    
    if 'pv' in name_lower or 'solar' in name_lower or '태양' in name_lower:
        return 'solar'
    elif 'wind' in name_lower or 'wt' in name_lower or '풍력' in name_lower:
        return 'wind'
    elif 'nuclear' in name_lower or '원자력' in name_lower:
        return 'nuclear'
    elif 'coal' in name_lower or '석탄' in name_lower:
        return 'coal'
    elif 'gas' in name_lower or '가스' in name_lower:
        return 'gas'
    elif 'oil' in name_lower or '석유' in name_lower:
        return 'oil'
    elif 'hydro' in name_lower or '수력' in name_lower:
        return 'hydro'
    elif 'biomass' in name_lower or '바이오' in name_lower:
        return 'biomass'
    elif 'battery' in name_lower or '배터리' in name_lower:
        return 'battery'
    elif 'hydrogen' in name_lower or 'h2' in name_lower or '수소' in name_lower:
        return 'hydrogen'
    else:
        return 'unknown'


def analyze_renewable_capacity_factor(gen_output, gen_info=None):
    """재생에너지 발전기의 용량 사용률(Capacity Factor)을 분석합니다."""
    if gen_output is None:
        print("발전기 출력 데이터가 없어 용량 사용률을 계산할 수 없습니다.")
        return None
    
    # 발전기 정보가 없는 경우 이름에서 유형 추론
    renewable_carriers = ['solar', 'wind', 'hydro', 'biomass']
    renewable_gens = []
    
    if gen_info is not None and 'carrier' in gen_info.columns:
        for idx, row in gen_info.iterrows():
            carrier = row.get('carrier', '').lower()
            name = row.get('name', '')
            
            # 발전기 정보에서 캐리어가 있는 경우
            if carrier in renewable_carriers:
                renewable_gens.append(name)
            # 캐리어 정보가 없는 경우 이름에서 추론
            elif infer_carrier_from_name(name) in renewable_carriers:
                renewable_gens.append(name)
    else:
        # 발전기 정보가 없는 경우 이름에서 유형 추론
        for col in gen_output.columns:
            if infer_carrier_from_name(col) in renewable_carriers:
                renewable_gens.append(col)
    
    if not renewable_gens:
        print("재생에너지 발전기를 찾을 수 없습니다.")
        return None
    
    print(f"{len(renewable_gens)}개의 재생에너지 발전기를 찾았습니다.")
    
    # 발전기 용량 정보 확인
    gen_capacity = {}
    if gen_info is not None and 'p_nom' in gen_info.columns and 'name' in gen_info.columns:
        for _, row in gen_info.iterrows():
            if row['name'] in gen_output.columns:
                gen_capacity[row['name']] = row['p_nom']
    
    if not gen_capacity:
        print("발전기 용량 정보가 없어 정확한 용량 사용률을 계산할 수 없습니다.")
        print("최대 출력값을 기준으로 상대적인 용량 사용률을 계산합니다.")
    
    # 재생에너지 발전기별 용량 사용률 계산
    capacity_factors = {}
    total_renewable_output = 0
    
    for gen in renewable_gens:
        if gen in gen_output.columns:
            # 발전량 계산
            output = gen_output[gen]
            total_output = output.sum()
            total_renewable_output += total_output
            
            # 발전기 용량
            if gen in gen_capacity:
                capacity = gen_capacity[gen]
                # 용량 사용률 = 발전량 / (용량 * 시간)
                time_hours = len(output)  # 시간 단위로 가정
                max_possible_output = capacity * time_hours
                cf = total_output / max_possible_output if max_possible_output > 0 else 0
                capacity_factors[gen] = {
                    'capacity': capacity,
                    'total_output': total_output,
                    'capacity_factor': cf,
                    'avg_output': output.mean(),
                    'max_output': output.max(),
                    'carrier': infer_carrier_from_name(gen)
                }
            else:
                # 용량 정보가 없는 경우 최대 출력을 용량으로 가정
                max_output = output.max()
                avg_output = output.mean()
                relative_cf = avg_output / max_output if max_output > 0 else 0
                capacity_factors[gen] = {
                    'capacity': '알 수 없음',
                    'total_output': total_output,
                    'relative_capacity_factor': relative_cf,
                    'avg_output': avg_output,
                    'max_output': max_output,
                    'carrier': infer_carrier_from_name(gen)
                }
    
    # 결과 요약
    print(f"\n재생에너지 발전 비중 분석 결과:")
    print(f"- 총 재생에너지 발전량: {total_renewable_output:.2f} MWh")
    
    # 발전기별 용량 사용률 요약
    carrier_output = {}
    carrier_count = {}
    
    for gen, data in capacity_factors.items():
        carrier = data['carrier']
        if carrier not in carrier_output:
            carrier_output[carrier] = 0
            carrier_count[carrier] = 0
        
        carrier_output[carrier] += data['total_output']
        carrier_count[carrier] += 1
    
    print("\n에너지원별 발전량:")
    for carrier, output in carrier_output.items():
        print(f"- {carrier}: {output:.2f} MWh ({carrier_count[carrier]}개 발전기)")
    
    return capacity_factors


def analyze_renewable_patterns(gen_output, gen_info=None):
    """재생에너지 발전기의 시간별 패턴을 분석합니다."""
    if gen_output is None or gen_output.empty:
        print("발전기 출력 데이터가 없어 시간별 패턴을 분석할 수 없습니다.")
        return
    
    # 재생에너지 발전기 식별
    renewable_carriers = ['solar', 'wind']
    renewable_gens = {}
    
    # 발전기 이름으로 유형 추론
    for col in gen_output.columns:
        carrier = infer_carrier_from_name(col)
        if carrier in renewable_carriers:
            if carrier not in renewable_gens:
                renewable_gens[carrier] = []
            renewable_gens[carrier].append(col)
    
    if not renewable_gens:
        print("재생에너지 발전기를 찾을 수 없습니다.")
        return
    
    print("\n재생에너지 시간별 패턴 분석:")
    for carrier, gens in renewable_gens.items():
        print(f"\n{carrier.upper()} 발전기 ({len(gens)}개):")
        
        if not gens:
            continue
        
        # 시간별 출력 합계
        carrier_output = gen_output[gens].sum(axis=1)
        
        # 일별/시간별 평균 계산
        if isinstance(carrier_output.index, pd.DatetimeIndex):
            try:
                # 시간별 평균 출력
                hourly_avg = carrier_output.groupby(carrier_output.index.hour).mean()
                
                # 일별 평균 출력
                daily_avg = carrier_output.groupby(carrier_output.index.date).mean()
                
                print(f"- 평균 출력: {carrier_output.mean():.2f} MW")
                print(f"- 최대 출력: {carrier_output.max():.2f} MW")
                print(f"- 최소 출력: {carrier_output.min():.2f} MW")
                
                # 일변화 패턴 시각화
                plt.figure(figsize=(12, 6))
                
                # 시간별 패턴
                plt.subplot(1, 2, 1)
                plt.plot(hourly_avg.index, hourly_avg.values)
                plt.title(f'{carrier.upper()} 시간별 평균 출력 패턴')
                plt.xlabel('시간 (0-23)')
                plt.ylabel('평균 출력 (MW)')
                plt.grid(True)
                
                # 일별 패턴
                plt.subplot(1, 2, 2)
                plt.plot(range(len(daily_avg)), daily_avg.values)
                plt.title(f'{carrier.upper()} 일별 평균 출력 패턴')
                plt.xlabel('일수')
                plt.ylabel('평균 출력 (MW)')
                plt.grid(True)
                
                # 그래프 저장
                save_dir = os.path.join('analysis_results')
                os.makedirs(save_dir, exist_ok=True)
                plt.tight_layout()
                plt.savefig(os.path.join(save_dir, f'{carrier}_pattern.png'))
                plt.close()
                
                print(f"- 패턴 그래프가 {save_dir}/{carrier}_pattern.png에 저장되었습니다.")
                
                # 시간별 출력 저장
                carrier_output.to_csv(os.path.join(save_dir, f'{carrier}_hourly_output.csv'))
                hourly_avg.to_csv(os.path.join(save_dir, f'{carrier}_hourly_avg.csv'))
                daily_avg.to_csv(os.path.join(save_dir, f'{carrier}_daily_avg.csv'))
                print(f"- 시간별 출력이 {save_dir}/{carrier}_hourly_output.csv에 저장되었습니다.")
                
            except Exception as e:
                print(f"시간별 패턴 분석 중 오류 발생: {str(e)}")
        else:
            print("출력 데이터의 인덱스가 시간 형식이 아니므로 시간별 패턴을 분석할 수 없습니다.")


def analyze_expansion_results(gen_output, gen_info=None):
    """발전기 확장 결과를 분석합니다."""
    if gen_output is None or gen_info is None:
        print("발전기 출력 또는 정보 데이터가 없어 확장 결과를 분석할 수 없습니다.")
        return None
    
    # 가상 발전기 및 재생에너지 발전기 식별
    virtual_gens = []
    renewable_gens = []
    conventional_gens = []
    
    for idx, gen in gen_info.iterrows():
        gen_name = gen.get('name', '')
        if gen_name not in gen_output.columns:
            continue
            
        carrier = gen.get('carrier', '').lower()
        gen_name_lower = gen_name.lower()
        
        # 가상 발전기 확인
        if gen_name.startswith('Virt_') or 'virtual' in carrier:
            virtual_gens.append(gen_name)
        # 재생에너지 발전기 확인
        elif ('pv' in gen_name_lower or 'solar' in gen_name_lower or 'solar' in carrier or '태양' in gen_name_lower or
              'wt' in gen_name_lower or 'wind' in gen_name_lower or 'wind' in carrier or '풍력' in gen_name_lower):
            renewable_gens.append(gen_name)
        else:
            conventional_gens.append(gen_name)
    
    print(f"\n발전기 확장 결과 분석:")
    print(f"- 가상 발전기: {len(virtual_gens)}개")
    print(f"- 재생에너지 발전기: {len(renewable_gens)}개")
    print(f"- 기타 발전기: {len(conventional_gens)}개")
    
    # 발전량 분석
    if gen_output is not None:
        # 가상 발전기 사용량
        virtual_output = gen_output[virtual_gens].sum().sum() if virtual_gens else 0
        renewable_output = gen_output[renewable_gens].sum().sum() if renewable_gens else 0
        conventional_output = gen_output[conventional_gens].sum().sum() if conventional_gens else 0
        total_output = gen_output.sum().sum()
        
        print(f"\n발전량 분석:")
        print(f"- 총 발전량: {total_output:.2f} MWh")
        print(f"- 가상 발전기 발전량: {virtual_output:.2f} MWh ({(virtual_output/total_output*100):.2f}%)")
        print(f"- 재생에너지 발전량: {renewable_output:.2f} MWh ({(renewable_output/total_output*100):.2f}%)")
        print(f"- 기타 발전기 발전량: {conventional_output:.2f} MWh ({(conventional_output/total_output*100):.2f}%)")
        
        # 가상 발전기 상위 사용량
        if virtual_gens and virtual_output > 0:
            virtual_outputs = {}
            for gen in virtual_gens:
                if gen in gen_output.columns:
                    virtual_outputs[gen] = gen_output[gen].sum()
            
            virtual_outputs = sorted(virtual_outputs.items(), key=lambda x: x[1], reverse=True)
            print("\n가상 발전기 상위 사용량:")
            for i, (gen, output) in enumerate(virtual_outputs[:10]):  # 상위 10개만
                if output > 0:
                    print(f"  {i+1}. {gen}: {output:.2f} MWh ({(output/virtual_output*100):.2f}%)")
        
        # 재생에너지 발전기 상위 사용량
        if renewable_gens and renewable_output > 0:
            renewable_outputs = {}
            for gen in renewable_gens:
                if gen in gen_output.columns:
                    renewable_outputs[gen] = gen_output[gen].sum()
            
            renewable_outputs = sorted(renewable_outputs.items(), key=lambda x: x[1], reverse=True)
            print("\n재생에너지 발전기 상위 사용량:")
            for i, (gen, output) in enumerate(renewable_outputs[:10]):  # 상위 10개만
                print(f"  {i+1}. {gen}: {output:.2f} MWh ({(output/renewable_output*100):.2f}%)")
    
    # 확장 결과 분석 (p_nom_opt - p_nom)
    if gen_info is not None and 'p_nom' in gen_info.columns and 'p_nom_opt' in gen_info.columns:
        # 확장된 발전기 수
        expanded_gens = gen_info[gen_info['p_nom_opt'] > gen_info['p_nom']]
        renewable_expanded = []
        virtual_expanded = []
        conventional_expanded = []
        
        for idx, gen in expanded_gens.iterrows():
            gen_name = gen.get('name', '')
            carrier = gen.get('carrier', '').lower()
            gen_name_lower = gen_name.lower()
            
            # 가상 발전기 확인
            if gen_name.startswith('Virt_') or 'virtual' in carrier:
                virtual_expanded.append(gen_name)
            # 재생에너지 발전기 확인
            elif ('pv' in gen_name_lower or 'solar' in gen_name_lower or 'solar' in carrier or '태양' in gen_name_lower or
                  'wt' in gen_name_lower or 'wind' in gen_name_lower or 'wind' in carrier or '풍력' in gen_name_lower):
                renewable_expanded.append(gen_name)
            else:
                conventional_expanded.append(gen_name)
        
        print(f"\n확장된 발전기 분석:")
        print(f"- 총 확장된 발전기: {len(expanded_gens)}개 중")
        print(f"  - 가상 발전기: {len(virtual_expanded)}개")
        print(f"  - 재생에너지 발전기: {len(renewable_expanded)}개")
        print(f"  - 기타 발전기: {len(conventional_expanded)}개")
        
        # 확장 규모 분석
        renewable_expansion = 0
        virtual_expansion = 0
        conventional_expansion = 0
        
        for idx, gen in expanded_gens.iterrows():
            gen_name = gen.get('name', '')
            p_nom = gen.get('p_nom', 0)
            p_nom_opt = gen.get('p_nom_opt', 0)
            expansion = p_nom_opt - p_nom
            
            if gen_name in virtual_gens:
                virtual_expansion += expansion
            elif gen_name in renewable_gens:
                renewable_expansion += expansion
            else:
                conventional_expansion += expansion
        
        total_expansion = renewable_expansion + virtual_expansion + conventional_expansion
        
        print(f"\n확장 규모 분석:")
        print(f"- 총 확장 용량: {total_expansion:.2f} MW")
        if total_expansion > 0:
            print(f"  - 가상 발전기 확장: {virtual_expansion:.2f} MW ({(virtual_expansion/total_expansion*100):.2f}%)")
            print(f"  - 재생에너지 확장: {renewable_expansion:.2f} MW ({(renewable_expansion/total_expansion*100):.2f}%)")
            print(f"  - 기타 발전기 확장: {conventional_expansion:.2f} MW ({(conventional_expansion/total_expansion*100):.2f}%)")
        
        # 상위 확장된 발전기 출력
        if len(expanded_gens) > 0:
            expansion_data = []
            for idx, gen in expanded_gens.iterrows():
                gen_name = gen.get('name', '')
                p_nom = gen.get('p_nom', 0)
                p_nom_opt = gen.get('p_nom_opt', 0)
                expansion = p_nom_opt - p_nom
                expansion_ratio = expansion / p_nom if p_nom > 0 else float('inf')
                
                gen_type = "가상" if gen_name in virtual_gens else "재생" if gen_name in renewable_gens else "기타"
                
                expansion_data.append({
                    'name': gen_name,
                    'type': gen_type,
                    'p_nom': p_nom,
                    'p_nom_opt': p_nom_opt,
                    'expansion': expansion,
                    'expansion_ratio': expansion_ratio
                })
            
            # 확장 규모 기준으로 정렬
            expansion_data.sort(key=lambda x: x['expansion'], reverse=True)
            
            print("\n상위 확장된 발전기(확장 규모 기준):")
            for i, data in enumerate(expansion_data[:10]):  # 상위 10개만
                print(f"  {i+1}. {data['name']} ({data['type']}): {data['p_nom']:.2f} MW → {data['p_nom_opt']:.2f} MW " +
                      f"(+{data['expansion']:.2f} MW, {(data['expansion_ratio']*100):.2f}% 증가)")
            
            # 확장 비율 기준으로 정렬
            expansion_data.sort(key=lambda x: x['expansion_ratio'], reverse=True)
            
            print("\n상위 확장된 발전기(확장 비율 기준):")
            for i, data in enumerate(expansion_data[:10]):  # 상위 10개만
                if data['p_nom'] > 0:  # 0으로 나누기 방지
                    print(f"  {i+1}. {data['name']} ({data['type']}): {data['p_nom']:.2f} MW → {data['p_nom_opt']:.2f} MW " +
                          f"(+{data['expansion']:.2f} MW, {(data['expansion_ratio']*100):.2f}% 증가)")
    
    return {
        'virtual_gens': virtual_gens,
        'renewable_gens': renewable_gens,
        'conventional_gens': conventional_gens
    }


def analyze_generator_results():
    """발전기 출력 결과를 분석합니다."""
    # 최신 결과 디렉토리 찾기
    results_dir = find_latest_results()
    if not results_dir:
        print("분석을 중단합니다.")
        return
    
    # 발전기 출력 데이터 로드
    gen_output = load_generator_output(results_dir)
    if gen_output is None:
        print("발전기 출력 데이터를 로드할 수 없어 분석을 중단합니다.")
        return
    
    # 발전기 정보 로드
    gen_info = load_generator_info(results_dir)
    
    # 분석 결과 저장 디렉토리 생성
    save_dir = os.path.join('analysis_results')
    os.makedirs(save_dir, exist_ok=True)
    
    print(f"\n분석 중... (발전기: {gen_output.shape[1]}개, 시간 단계: {gen_output.shape[0]}개)")
    
    # 발전량 총합 계산
    total_generation = gen_output.sum().sum()
    print(f"총 발전량: {total_generation:.2f} MWh")
    
    # 발전기별 에너지원 추론
    carrier_output = {}
    for col in gen_output.columns:
        carrier = infer_carrier_from_name(col)
        if carrier not in carrier_output:
            carrier_output[carrier] = 0
        carrier_output[carrier] += gen_output[col].sum()
    
    # 에너지원별 발전량 저장
    carrier_df = pd.DataFrame({
        '에너지원': list(carrier_output.keys()),
        '발전량_MWh': [carrier_output[c] for c in carrier_output.keys()],
        '비율': [carrier_output[c]/total_generation*100 for c in carrier_output.keys()]
    })
    carrier_df = carrier_df.sort_values('발전량_MWh', ascending=False)
    carrier_df.to_csv(os.path.join(save_dir, 'generation_by_carrier.csv'), index=False, encoding='utf-8-sig')
    print(f"에너지원별 발전량이 {save_dir}/generation_by_carrier.csv에 저장되었습니다.")
    
    # 에너지원별 발전량 출력
    print("\n에너지원별 발전량:")
    for idx, row in carrier_df.iterrows():
        print(f"- {row['에너지원']}: {row['발전량_MWh']:.2f} MWh ({row['비율']:.1f}%)")
    
    # 지역별 발전량 분석 - 발전기 이름에서 지역 추출
    region_pattern = r'([A-Za-z가-힣]+)_'  # 지역명_나머지 패턴 가정
    region_output = {}
    
    for col in gen_output.columns:
        match = re.match(region_pattern, col)
        region = match.group(1) if match else "기타"
        
        if region not in region_output:
            region_output[region] = 0
        region_output[region] += gen_output[col].sum()
    
    # 지역별 발전량 저장
    region_df = pd.DataFrame({
        '지역': list(region_output.keys()),
        '발전량_MWh': [region_output[r] for r in region_output.keys()],
        '비율': [region_output[r]/total_generation*100 for r in region_output.keys()]
    })
    region_df = region_df.sort_values('발전량_MWh', ascending=False)
    region_df.to_csv(os.path.join(save_dir, 'generation_by_region.csv'), index=False, encoding='utf-8-sig')
    print(f"지역별 발전량이 {save_dir}/generation_by_region.csv에 저장되었습니다.")
    
    # 지역별 발전량 출력
    print("\n지역별 발전량:")
    for idx, row in region_df.iterrows():
        print(f"- {row['지역']}: {row['발전량_MWh']:.2f} MWh ({row['비율']:.1f}%)")
    
    # 발전기 확장 결과 분석
    gen_categories = analyze_expansion_results(gen_output, gen_info)
    
    # 재생에너지 발전기의 용량 사용률 분석
    capacity_factors = analyze_renewable_capacity_factor(gen_output, gen_info)
    
    # 재생에너지 시간별 패턴 분석
    analyze_renewable_patterns(gen_output, gen_info)
    
    # 결과 요약 Excel 파일 생성
    with pd.ExcelWriter(os.path.join(save_dir, 'generation_summary.xlsx')) as writer:
        carrier_df.to_excel(writer, sheet_name='에너지원별', index=False)
        region_df.to_excel(writer, sheet_name='지역별', index=False)
        
        # 발전기별 상세 정보 시트 추가
        gen_detail = pd.DataFrame({
            '발전기': gen_output.columns,
            '총발전량_MWh': [gen_output[col].sum() for col in gen_output.columns],
            '평균발전량_MW': [gen_output[col].mean() for col in gen_output.columns],
            '최대발전량_MW': [gen_output[col].max() for col in gen_output.columns],
            '사용시간': [(gen_output[col] > 0).sum() for col in gen_output.columns],
            '에너지원': [infer_carrier_from_name(col) for col in gen_output.columns]
        })
        gen_detail = gen_detail.sort_values('총발전량_MWh', ascending=False)
        gen_detail.to_excel(writer, sheet_name='발전기별', index=False)
        
        # 가상 발전기와 재생에너지 발전기 구분해서 시트 추가 (정보가 있는 경우)
        if gen_categories and gen_info is not None:
            # 발전기 정보와 출력 통합
            if 'name' in gen_info.columns:
                gen_full_info = pd.merge(
                    gen_info, 
                    pd.DataFrame({
                        'name': gen_output.columns,
                        'total_output': [gen_output[col].sum() for col in gen_output.columns],
                        'avg_output': [gen_output[col].mean() for col in gen_output.columns],
                        'max_output': [gen_output[col].max() for col in gen_output.columns],
                        'usage_hours': [(gen_output[col] > 0).sum() for col in gen_output.columns]
                    }),
                    on='name', how='inner'
                )
                
                # 확장률 계산
                if 'p_nom' in gen_full_info.columns and 'p_nom_opt' in gen_full_info.columns:
                    gen_full_info['expansion'] = gen_full_info['p_nom_opt'] - gen_full_info['p_nom']
                    gen_full_info['expansion_ratio'] = gen_full_info['expansion'] / gen_full_info['p_nom'].replace(0, float('nan'))
                
                # 가상 발전기 정보
                virtual_info = gen_full_info[gen_full_info['name'].isin(gen_categories['virtual_gens'])]
                if not virtual_info.empty:
                    virtual_info = virtual_info.sort_values('total_output', ascending=False)
                    virtual_info.to_excel(writer, sheet_name='가상발전기', index=False)
                
                # 재생에너지 발전기 정보
                renewable_info = gen_full_info[gen_full_info['name'].isin(gen_categories['renewable_gens'])]
                if not renewable_info.empty:
                    renewable_info = renewable_info.sort_values('total_output', ascending=False)
                    renewable_info.to_excel(writer, sheet_name='재생에너지', index=False)
                
                # 확장된 발전기 정보
                if 'expansion' in gen_full_info.columns:
                    expanded_info = gen_full_info[gen_full_info['expansion'] > 0]
                    if not expanded_info.empty:
                        expanded_info = expanded_info.sort_values('expansion', ascending=False)
                        expanded_info.to_excel(writer, sheet_name='확장발전기', index=False)
    
    print(f"종합 분석 결과가 {save_dir}/generation_summary.xlsx에 저장되었습니다.")
    
    # 에너지원별 발전량 시각화
    plt.figure(figsize=(12, 10))
    
    # 원형 차트 - 에너지원별
    plt.subplot(2, 1, 1)
    plt.pie(carrier_df['발전량_MWh'], labels=carrier_df['에너지원'], 
            autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('에너지원별 발전량 비중')
    
    # 막대 그래프 - 지역별
    plt.subplot(2, 1, 2)
    plt.bar(region_df['지역'], region_df['발전량_MWh'])
    plt.title('지역별 발전량')
    plt.xlabel('지역')
    plt.ylabel('발전량 (MWh)')
    plt.xticks(rotation=45)
    
    plt.tight_layout()
    plt.savefig(os.path.join(save_dir, 'generation_charts.png'))
    plt.close()
    
    print(f"발전량 차트가 {save_dir}/generation_charts.png에 저장되었습니다.")
    
    return {
        'total_generation': total_generation,
        'by_carrier': carrier_df,
        'by_region': region_df,
        'save_dir': save_dir
    }


if __name__ == "__main__":
    results = analyze_generator_results()
    if results:
        print("\n분석이 성공적으로 완료되었습니다.")
        print(f"총 발전량: {results['total_generation']:.2f} MWh")
        print(f"상위 3개 에너지원: {', '.join(results['by_carrier']['에너지원'].head(3).tolist())}")
        print(f"상위 3개 지역: {', '.join(results['by_region']['지역'].head(3).tolist())}")
        print(f"분석 결과 파일: {results['save_dir']}")
    else:
        print("\n분석을 완료할 수 없습니다.") 