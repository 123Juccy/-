import numpy as np
import soundfile as sf
import os

def attenuate_audio(input_path, output_path, attenuation_db=50):
    """
    读取WAV文件，衰减50dB，保存为32位浮点WAV
    
    参数:
        input_path: 输入WAV文件路径
        output_path: 输出WAV文件路径
        attenuation_db: 衰减分贝数(默认50dB)
    """
    try:
        # 读取音频文件
        audio, sample_rate = sf.read(input_path)
        
        # 确保音频是单声道(如果是立体声则取平均值)
        if audio.ndim > 1:
            audio = np.mean(audio, axis=1)
        
        # 计算衰减比例(线性值)
        attenuation_linear = 10 ** (-attenuation_db / 20)
        
        # 应用衰减
        attenuated_audio = audio * attenuation_linear
        
        # 确保数据在32位浮点范围内
        attenuated_audio = np.clip(attenuated_audio, -1.0, 1.0).astype(np.float32)
        
        # 创建输出目录(如果不存在)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 保存衰减后的音频(32位浮点)
        sf.write(output_path, attenuated_audio, sample_rate, subtype='FLOAT')
        
        print(f"处理成功: {input_path} -> {output_path}")
        
    except Exception as e:
        print(f"处理失败 {input_path}: {str(e)}")

def process_directory(input_dir, output_dir, attenuation_db=50):
    """
    递归处理目录中的所有WAV文件
    
    参数:
        input_dir: 输入目录路径
        output_dir: 输出目录路径
        attenuation_db: 衰减分贝数(默认50dB)
    """
    for root, dirs, files in os.walk(input_dir):
        # 在输出目录中保持相同的子目录结构
        relative_path = os.path.relpath(root, input_dir)
        current_output_dir = os.path.join(output_dir, relative_path)
        
        for file in files:
            if file.lower().endswith('.wav'):
                input_path = os.path.join(root, file)
                output_path = os.path.join(current_output_dir, file)
                
                # 处理音频文件
                attenuate_audio(input_path, output_path, attenuation_db)

if __name__ == "__main__":
    input_directory = input("请输入包含WAV文件的目录路径: ").strip('"')
    output_directory = input("请输入输出目录路径: ").strip('"')
    
    # 检查输入目录是否存在
    if not os.path.isdir(input_directory):
        print(f"错误: 输入目录不存在 {input_directory}")
        exit(1)
    
    # 执行批量处理
    print("开始批量处理WAV文件...")
    process_directory(input_directory, output_directory)
    print("批量处理完成!")