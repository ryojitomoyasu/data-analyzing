print("=== test_ai_engine.py started ===")

from ai_engine import AIEngine

try:
    engine = AIEngine()

    result = engine.get_strategy_advice(
        user_query="売上が伸び悩んでいます。どう改善すべき？",
        internal_data_summary="売上前年比-10%、来店数横ばい、単価減少",
        external_trends="同業他社はD2CとSNS施策を強化"
    )

    print("=== result type ===", type(result))
    print("=== result repr ===", repr(result))
    print("=== result ===")
    print(result)

except Exception as e:
    print("=== ERROR OCCURRED ===")
    import traceback
    traceback.print_exc()
