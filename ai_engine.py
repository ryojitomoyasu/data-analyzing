import os
try:
    from google import genai
except ImportError:
    genai = None

from dotenv import load_dotenv

load_dotenv()

class AIEngine:
    def __init__(self, api_key=None):
        key = api_key or os.getenv("GOOGLE_API_KEY")
        if not key:
            raise ValueError("Google API Key is required.")
        
        if genai is None:
            raise ImportError("Google GenAI library is not installed. Please install 'google-genai'.")
        
        # Initialize the new GenAI client
        self.client = genai.Client(api_key=key)
        # Use Gemini 2.0 Flash as requested
        self.model_name = "gemini-2.0-flash"

    def get_industry_trends(self, industry_context):
        """
        Fetch industry trends.
        """
        prompt = f"""
以下の業界、または商品群に関する2025年最新のトレンドと市場の課題を調査してください。
経営戦略に役立つ具体的な情報を要約してください。

対象: {industry_context}
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt
            )
            return response.text
        except Exception as e:
            return f"Error fetching trends: {e}"

    def infer_industry_smart(self, sample_data_text):
        """
        Infer industry using Gemini based on sample data.
        """
        prompt = f"""
提供された商品名のリストから、このビジネスが属する「業界名」を短く（10文字以内）で出力してください。

商品リスト:
{sample_data_text}

業界名のみを出力してください。
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt
            )
            return response.text.strip()
        except Exception as e:
            return "General"

    def get_strategy_advice(self, user_query, internal_data_summary, external_trends, history=None):
        """
        Chat with a strategy consultant persona.
        """
        system_prompt = f"""
あなたは世界トップクラスの戦略コンサルティングファームのシニアパートナーです。
MBAレベルの高度な経営理論（マーケティング4P、アンゾフの成長マトリクス、SWOT分析、パレートの法則など）を駆使し、クライアント（ユーザー）のデータに基づいて鋭い洞察と具体的な戦略提案を行ってください。

【あなたの行動指針】
1. **Facts First**: 必ず提示された【内部データ統計】の数値を根拠に話を展開してください。「売上が上がっています」ではなく「昨対比+15%の成長が見られますが、これはX商品が牽引しています」のように具体的に述べてください。
2. **"So What?"**: 単なるデータの羅列ではなく、「だから何なのか？（So What?）」という示唆、そして「どうすべきか？（Action）」まで踏み込んでください。
3. **Hypothesis Driven**: データから読み取れる仮説を立て、それを検証するための視点を提供してください。
4. **Tone**: プロフェッショナルかつ論理的、しかしクライアントに寄り添うパートナーとして振る舞ってください。

【ユーザーの内部データ統計】
{internal_data_summary}

【外部トレンド情報】
{external_trends}

上記の情報に基づき、ユーザーの質問に回答してください。
"""
        
        # Simple History handling
        # Optimization: Truncate history to last 6 messages to save tokens/context
        MAX_HISTORY = 6
        history_text = ""
        
        if history:
            # Take only the last N messages
            recent_history = history[-MAX_HISTORY:] if len(history) > MAX_HISTORY else history
            for msg in recent_history:
                role = "User" if msg["role"] == "user" else "Assistant"
                # Truncate very long messages individually too (e.g. 500 chars)
                content = msg['content']
                if len(content) > 500:
                    content = content[:500] + "...(truncated)"
                history_text += f"{role}: {content}\n"
        
        full_prompt = f"""{system_prompt}

【これまでの会話】
{history_text}

【ユーザーの質問】
{user_query}
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=full_prompt
            )
            return response.text
        except Exception as e:
            return f"Error getting advice: {e}"

    def generate_sales_insight(self, data_summary):
        """
        Generate a concise 3-line insight based on dashboard data.
        """
        prompt = f"""
以下の売上データを分析し、経営者向けに「好調/不調な要因」と「次に取り組むべき重要ポイント」を簡潔にまとめてください。
必ず **3行以内** の箇条書きで出力してください。

【データ概要】
{data_summary}

出力形式:
- [Good] ...
- [Bad/Risk] ...
- [Action] ...
"""
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt
            )
            return response.text
        except Exception as e:
            return f"Error generating insight: {str(e)}"

