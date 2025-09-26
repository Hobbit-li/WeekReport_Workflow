import pandas as pd
from openai import OpenAI
import chromadb  # 作为知识库
from typing import Any, Dict

from Excel_Convert import generate_weekly_report

# --------------------
# 构建知识库
# --------------------
class KnowledgeBase:
    def __init__(self, persist_dir: str = "./chroma_store"):
        self.client = chromadb.PersistentClient(path=persist_dir)
        self.collection = self.client.get_or_create_collection("analysis_reports")

    def add_report(self, report_id: str, text: str, metadata: Dict[str, Any] = None):
        self.collection.add(documents=[text], ids=[report_id], metadatas=[metadata])

    def query(self, query_text: str, top_k: int = 3) -> list[str]:
        results = self.collection.query(query_texts=[query_text], n_results=top_k)
        return results["documents"][0]

# --------------------
# 调用大语言模型
# --------------------
class LLMAnalyzer:
    def __init__(self, api_key: str):
        self.client = OpenAI(api_key=api_key)

    def analyze(self, table_summary: str, kb_context: str = "") -> str:
        prompt = f"""
            你是一个数据分析专家。下面是数据表的摘要：
            {table_summary}
            
            以下是过往的参考分析：
            {kb_context}
            
            请生成一份包含表格解读、关键发现和结论的分析报告，报告要结构化且简洁。
            """
        response = self.client.chat.completions.create(
            model="gpt-4o-mini",  # 可以换成你有权限的模型
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content

# --------------------
# Step 5: 主工作流
# --------------------
def main(file_path: str, api_key: str):
    # 1. 读取数据
    raw_df = load_raw_data(file_path)

    # 2. 预处理
    processed_df = generate_weekly_report(raw_df)

    # 3. 获取表格摘要（可截取前几行 + 描述统计）
    summary = processed_df.head(10).to_markdown() + "\n\n" + str(processed_df.describe())

    # 4. 知识库检索
    kb = KnowledgeBase()
    kb_context = "\n".join(kb.query("数据分析报告"))

    # 5. 调用 LLM 分析
    analyzer = LLMAnalyzer(api_key)
    report = analyzer.analyze(summary, kb_context)

    # 6. 存储报告到知识库
    kb.add_report(report_id="report_001", text=report, metadata={"file": file_path})

    # 7. 输出报告
    with open("analysis_report.md", "w", encoding="utf-8") as f:
        f.write(report)

    print("分析报告已生成：analysis_report.md")

# --------------------
# 执行
# --------------------
if __name__ == "__main__":
    main("data.csv", api_key="your_api_key_here")
