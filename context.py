class Context:
    automacao_fusion_instance = None

def set_automacao_fusion_instance(instance):
    Context.automacao_fusion_instance = instance

def get_automacao_fusion_instance():
    return Context.automacao_fusion_instance