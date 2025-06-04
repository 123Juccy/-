import os
import librosa
import soundfile as sf
import numpy as np

# 输入输出文件夹路径
input_folder = 'C:/Users/OSS360211/Desktop/laifen1_2/ng'  # 请替换为你的音频文件所在文件夹路径
output_folder = 'C:/Users/OSS360211/Desktop/lf1_2new/ng'  # 请替换为你想保存分割后音频文件的文件夹路径
part_duration_ms = 10000  # 自定义每一部分的时长，单位为毫秒，例如 5000ms 即 5秒

# 如果输出文件夹不存在，则创建它
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取输入文件夹中所有音频文件
audio_files = [f for f in os.listdir(input_folder) if f.endswith('.wav') or f.endswith('.mp3')]  # 根据实际文件类型调整

# 遍历每个音频文件
for audio_file in audio_files:
    audio_path = os.path.join(input_folder, audio_file)

    # 使用 librosa 加载音频文件，返回音频数据和采样率
    audio_data, sample_rate = librosa.load(audio_path, sr=None)  # sr=None 保持原采样率

    # 获取音频的总时长（单位：样本数）
    total_samples = len(audio_data)

    # 将所需的分割时长转换为样本数
    part_duration_samples = (part_duration_ms / 1000) * sample_rate

    # 计算所需的分割数量
    num_parts = int(np.ceil(total_samples / part_duration_samples))

    # 分割音频并保存
    for i in range(num_parts):
        start_sample = int(i * part_duration_samples)
        end_sample = int(min((i + 1) * part_duration_samples, total_samples))

        # 获取当前部分的音频数据
        part = audio_data[start_sample:end_sample]

        # 如果是最后一部分且长度不足指定长度，使用静音填充
        if len(part) < part_duration_samples:
            silence_duration_samples = int(part_duration_samples - len(part))
            silence = np.zeros(silence_duration_samples)  # 创建静音段
            part = np.concatenate((part, silence))  # 在末尾添加静音

        # 保存分割后的音频文件
        output_file = os.path.join(output_folder, f'{os.path.splitext(audio_file)[0]}_part{i + 1}.wav')
        sf.write(output_file, part, sample_rate, subtype='PCM_32')  # 使用 soundfile 保存为 wav 格式

    print(f'{audio_file} 已分割为 {num_parts} 份并保存！')

print("所有音频文件已处理完毕！")