import pandas as pd
import os

input_file = 'output2/todos-telefones.xlsx'
df = pd.read_excel(input_file)

output_folder = 'output'
os.makedirs(output_folder, exist_ok=True)

max_lines = 999

num_chunks = len(df) // max_lines + (1 if len(df) % max_lines != 0 else 0)

for i in range(num_chunks):
    chunk = df.iloc[i * max_lines : (i + 1) * max_lines]
    output_file_csv = os.path.join(output_folder, f'todos_{i + 1:02d}.csv')
    chunk.to_csv(output_file_csv, index=False)
    print(f'Arquivo {output_file_csv} gerado com {len(chunk)} linhas.')

print("Divisão concluída.")
