class Context:
    automacao_fusion_instance = None

    @classmethod
    def set_automacao_fusion_instance(cls, instance):
        cls.automacao_fusion_instance = instance

    @classmethod
    def get_automacao_fusion_instance(cls):
        return cls.automacao_fusion_instance