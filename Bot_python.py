import logging
import requests
import json
import os
import io
import re
import random
import asyncio
import html
from dotenv import load_dotenv
from datetime import datetime, timedelta
from thefuzz import process

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, Message, CallbackQuery, Bot
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes,
    ConversationHandler,
)

# --- CONFIGURAÇÕES ---
load_dotenv()
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
APP_SCRIPT_URL = os.getenv("APP_SCRIPT_URL")
HUMANIZACAO_ATIVA = True  # Ativa ou desativa a humanização automática
if not TELEGRAM_TOKEN or not APP_SCRIPT_URL:
    raise ValueError(
        "As variáveis TELEGRAM_TOKEN e APP_SCRIPT_URL não foram definidas no arquivo .env"
    )

# Estados para a conversa
MOTIVO_REJEICAO = range(1)

# --- FUNÇÃO HELPER PARA CHAMAR A API ---
def call_app_script(action, params={}):
    """
    Chama a API do Google Apps Script e SEMPRE retorna um dicionário,
    garantindo que o bot nunca quebre por AttributeError.
    """
    try:
        logging.info(f"[API Call] Enviando ação: '{action}'")
        payload = json.dumps({"action": action, "params": params})
        headers = {"Content-Type": "application/json"}
        
        response = requests.post(
            APP_SCRIPT_URL, headers=headers, data=payload, timeout=90
        )
        
        # Lança um erro para status HTTP 4xx (ex: Not Found) ou 5xx (erro de servidor)
        response.raise_for_status() 
        
        # Tenta decodificar a resposta JSON. Se a resposta for um HTML de erro do Google,
        # isso vai falhar e pular para o bloco 'except json.JSONDecodeError'.
        json_response = response.json()
        
        logging.info(f"[API Response] Sucesso para ação: '{action}'")
        return json_response
        
    except requests.exceptions.Timeout:
        logging.error(f"[API Response] Timeout na ação '{action}'")
        return {"status": "timeout", "message": "O servidor demorou para responder."}
        
    except requests.exceptions.RequestException as e:
        # Pega outros erros de rede ou erros de status HTTP (4xx, 5xx)
        logging.error(f"[API Response] Erro de Rede/HTTP na ação '{action}': {e}")
        return {"status": "network_error", "message": f"Falha de comunicação: {e}"}
        
    except json.JSONDecodeError:
        # Isso acontece se o Apps Script quebrar e retornar uma página de erro HTML
        logging.error(f"[API Response] Erro de JSON na ação '{action}'. O servidor retornou uma resposta inválida.")
        return {"status": "invalid_response", "message": "O servidor retornou uma resposta inesperada."}

    except Exception as e:
        # Pega qualquer outro erro inesperado que possa acontecer
        logging.error(f"[API Response] Erro inesperado no call_app_script para '{action}': {e}")
        return {"status": "unknown_error", "message": f"Erro inesperado no Python: {e}"}

# Mantém referências aos métodos originais
_orig_message_reply_text = Message.reply_text
_orig_callback_edit_text = CallbackQuery.edit_message_text
_orig_bot_send_message = Bot.send_message
_orig_bot_send_document = Bot.send_document

# Adicione esta nova função perto das suas outras funções "helper"

async def build_paginated_keyboard(items: list, page: int = 0, items_per_page: int = 8) -> InlineKeyboardMarkup:
    """
    Cria um teclado InlineKeyboardMarkup paginado a partir de uma lista de itens.
    """
    start_index = page * items_per_page
    end_index = start_index + items_per_page
    
    # Pega apenas os itens da página atual
    page_items = items[start_index:end_index]
    
    botoes = []
    # Cria um botão para cada item da página
    for i, item_text in enumerate(page_items):
        # Calcula o índice real do item na lista completa
        item_index = start_index + i
        # AQUI ESTÁ A CORREÇÃO: Usamos o índice 'item_index' no callback_data
        botoes.append([InlineKeyboardButton(item_text, callback_data=f"buscar_fornecedor:{item_index}")])

    # Cria os botões de navegação
    navigation_row = []
    if page > 0:
        navigation_row.append(InlineKeyboardButton("◀️ Anterior", callback_data=f"page_fornecedores:{page - 1}"))
    
    if end_index < len(items):
        navigation_row.append(InlineKeyboardButton("Próximo ▶️", callback_data=f"page_fornecedores:{page + 1}"))
    
    if navigation_row:
        botoes.append(navigation_row)
        
    return InlineKeyboardMarkup(botoes)

def _choose_prefix_for(text: str) -> str:
    t = (text or "").strip()
    if not t or t[0] in "😃😀😅😂😉😊🙂🙃🤖📦✅❌🔎⏳":
        return ""
    low = t.lower()
    
    # MENSAGENS DE SISTEMA - mais conversacional
    if any(x in low for x in ("timeout", "servidor demorou", "demorou demais para responder")):
        return random.choice([
            "Poxa, parece que", "Eita, acho que", "Rapaz,", 
            "Que chato,", "Puxa vida,", "Ô, que demora..."
        ])
    if any(x in low for x in ("network_error", "falha de comunicação")):
        return random.choice([
            "Vixe, tivemos um problema aqui —", "Putz, algo deu errado —", 
            "Caramba, não consegui me conectar —", "Rapaz, a conexão falhou —",
            "Opa, deu zebra na conexão —"
        ])
    if any(x in low for x in ("invalid_response", "resposta inesperada")):
        return random.choice([
            "Hmm, recebi algo estranho do servidor —", "Opa, isso não era o que eu esperava —",
            "Eita, parece que o servidor enviou algo diferente —", "Rapaz, que resposta estranha..."
        ])
    
    # MENSAGENS DE ERRO GERAIS - tom compreensivo
    if any(x in low for x in ("erro", "falha", "não foi", "não consegui", "problema", "ocorreu um erro")):
        return random.choice(["Putz,", "Ops,", "Que pena,", "Vixe,", "Poxa,", "Eita,", "Nossa,"])
    
    # MENSAGENS DE SUCESSO - celebrativo
    if any(x in low for x in ("sucesso", "pronto", "concluído", "criado", "salvo", "gerado")):
        return random.choice(["Massa!", "Show!", "Perfeito!", "Beleza!", "Ótimo!", "Mandou bem!", "Feito!"])
    
    # MENSAGENS DE PROCESSAMENTO - tranquilizador
    if any(x in low for x in ("aguarde", "carregando", "sincronizando", "gerando", "buscando", "trabalhando", "processando")):
        return random.choice([
            "Calma aí que", "Só um minutinho,", "Deixa comigo,", 
            "Segura aí que", "Já estou trabalhando nisso,", "Peraí que"
        ])
    
    # CONFIRMAÇÕES E PERGUNTAS
    if any(x in low for x in ("confirma", "deseja", "qual", "como", "quando", "onde")):
        return random.choice(["Então,", "Beleza,", "Certo,", "Perfeito,", "Show,"])
    
    # INFORMAÇÕES E INSTRUÇÕES
    if any(x in low for x in ("use:", "exemplo:", "formato", "digite", "envie")):
        return random.choice(["Ó só,", "Olha,", "Veja bem,", "Então,", "É assim:"])
    
    return random.choice(["Então,", "Beleza,", "Bom,", "Certo,", "Opa,"])

def _humanize_text(text: str, parse_mode: str | None = None) -> str:
    try:
        # Se a humanização estiver desativada, retorna o texto original
        if not HUMANIZACAO_ATIVA:
            return text
            
        if not isinstance(text, str):
            return text
        text = text.strip()
        
        # Evita humanizar texto que já tem prefixos conversacionais
        primeiras_palavras = text.split()[:2]
        if primeiras_palavras:
            primeira = primeiras_palavras[0].lower().rstrip(',')
            palavras_ja_humanizadas = [
                'opa', 'eita', 'putz', 'vixe', 'massa', 'show', 'beleza', 'então',
                'olha', 'olá', 'oi', 'bom', 'boa', 'claro', 'certo'
            ]
            if primeira in palavras_ja_humanizadas:
                return text  # Já está humanizado
        
        # Casos especiais para mensagens muito comuns
        if "não entendi" in text.lower():
            variações = [
                "Opa, não consegui entender direito o que você quis dizer.",
                "Hmm, não captei essa. Pode tentar de novo?",
                "Desculpa, não processei isso direito.",
                "Eita, não entendi. Pode reformular?"
            ]
            return random.choice(variações)
        
        if "tente novamente" in text.lower():
            variações = [
                "Que tal tentar mais uma vez?",
                "Bora tentar de novo?",
                "Vamos tentar outra vez?",
                "Tenta aí de novo, vai!"
            ]
            return random.choice(variações)
        
        # Sistema normal de prefixos
        prefix = _choose_prefix_for(text)
        if prefix and not text.startswith(prefix):
            return f"{prefix} {text}"
        return text
    except Exception:
        return text

# Wrappers async para manter compatibilidade com as assinaturas originais
async def _human_message_reply_text(self, text, *args, **kwargs):
    parse_mode = kwargs.get("parse_mode", None)
    human = _humanize_text(text, parse_mode=parse_mode)
    return await _orig_message_reply_text(self, human, *args, **kwargs)

async def _human_callback_edit_text(self, text, *args, **kwargs):
    parse_mode = kwargs.get("parse_mode", None)
    human = _humanize_text(text, parse_mode=parse_mode)
    return await _orig_callback_edit_text(self, human, *args, **kwargs)

async def _human_bot_send_message(self, chat_id, text, *args, **kwargs):
    parse_mode = kwargs.get("parse_mode", None)
    human = _humanize_text(text, parse_mode=parse_mode)
    return await _orig_bot_send_message(self, chat_id, human, *args, **kwargs)

async def _human_bot_send_document(self, chat_id, document, *args, **kwargs):
    # humaniza apenas a caption se existir
    caption = kwargs.get("caption", None)
    parse_mode = kwargs.get("parse_mode", None)
    if caption:
        kwargs["caption"] = _humanize_text(caption, parse_mode=parse_mode)
    return await _orig_bot_send_document(self, chat_id, document, *args, **kwargs)

# Adicione helpers específicos:
async def send_user_reply(message_obj, text, **kwargs):
    """Envia resposta humanizada apenas para interações diretas do usuário."""
    parse_mode = kwargs.get("parse_mode", None)
    human_text = _humanize_text(text, parse_mode=parse_mode)
    return await message_obj.reply_text(human_text, **kwargs)

async def edit_user_message(query_obj, text, **kwargs):
    """Edita mensagem com tom humanizado apenas para interações do usuário."""
    parse_mode = kwargs.get("parse_mode", None)
    human_text = _humanize_text(text, parse_mode=parse_mode)
    return await query_obj.edit_message_text(human_text, **kwargs)

async def send_system_reply(message_obj, text, **kwargs):
    """Envia mensagem do sistema SEM humanização (timeouts, erros, etc.)."""
    return await message_obj.reply_text(text, **kwargs)

async def edit_system_message(query_obj, text, **kwargs):
    """Edita mensagem do sistema SEM humanização."""
    return await query_obj.edit_message_text(text, **kwargs)

# Aplica o monkeypatch para humanização automática
#Message.reply_text = _human_message_reply_text
#CallbackQuery.edit_message_text = _human_callback_edit_text
#Bot.send_message = _human_bot_send_message
#Bot.send_document = _human_bot_send_document

def log_interaction(update: Update | None, bot_response: str = None):
    """
    Registra a interação de um usuário ou a resposta do bot em um arquivo de log.
    Usa caminho absoluto (mesma pasta do script) e tolera objetos 'update' incompletos.
    """
    try:
        # Garante que o arquivo será criado na mesma pasta do script, independente do cwd
        base_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(base_dir, "log_completo.txt")

        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Valores seguros para campos do usuário/mensagem
        user_id = "unknown"
        user_name = "Desconhecido"
        user_message = "<sem mensagem>"

        if update:
            user = getattr(update, "effective_user", None)
            if user:
                user_id = getattr(user, "id", user_id)
                user_name = getattr(user, "first_name", user_name) or user_name

            # tenta obter texto da mensagem ou callback
            eff_msg = getattr(update, "effective_message", None)
            if eff_msg and getattr(eff_msg, "text", None):
                user_message = eff_msg.text
            else:
                cb = getattr(update, "callback_query", None)
                if cb and getattr(cb, "data", None):
                    user_message = f"[BOTÃO CLICADO: {cb.data}]"

        # Monta a linha do log
        if bot_response:
            cleaned_response = ' '.join(bot_response.splitlines())
            log_message = f"[{timestamp}] ID: {user_id} ({user_name}) <<< BOT: {cleaned_response}\n"
        else:
            log_message = f"[{timestamp}] ID: {user_id} ({user_name}) >>> USUÁRIO: {user_message}\n"

        # Grava com 'a' (append) - cria o arquivo se não existir
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(log_message)

    except Exception as e:
        logging.error(f"Falha ao escrever no log de interação: {e}")

def ensure_log_file():
    """Garante que o arquivo de log exista (cria se necessário)."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(base_dir, "log_completo.txt")
    try:
        # Abre em modo append — cria o arquivo se não existir.
        with open(log_path, "a", encoding="utf-8"):
            pass
    except Exception as e:
        logging.warning(f"Não foi possível criar/abrir o arquivo de log '{log_path}': {e}")

def normalizar_placa(placa: str) -> str:
    """Remove hífen e espaços, deixa maiúsculo."""
    return placa.replace("-", "").replace(" ", "").upper()

def validar_e_normalizar_placa(placa: str) -> str | None:
    """
    Valida a placa contra os padrões BR (AAA-1234) e Mercosul (AAA1B34).
    Se for válida, retorna a placa normalizada (ex: 'ABC1234').
    Se for inválida, retorna None.
    """
    if not placa:
        return None

    # Normaliza a placa para o teste: maiúsculas, sem espaços ou hífens
    placa_normalizada = placa.strip().upper().replace("-", "")
    
    # Padrão Regex para placa antiga (BR): 3 letras, 4 números
    padrao_br = re.compile(r"^[A-Z]{3}\d{4}$")
    
    # Padrão Regex para placa nova (Mercosul): 3 letras, 1 número, 1 letra, 2 números
    padrao_mercosul = re.compile(r"^[A-Z]{3}\d[A-Z]\d{2}$")

    # Verifica se a placa normalizada corresponde a um dos padrões
    match_br = padrao_br.match(placa_normalizada)
    match_mercosul = padrao_mercosul.match(placa_normalizada)

    if match_br or match_mercosul:
        # --- NOVA VALIDAÇÃO DE SEQUÊNCIAS REPETITIVAS ---
        
        # 1. Verifica se as 3 primeiras letras são todas iguais
        letras = placa_normalizada[0:3]
        if len(set(letras)) == 1:
            return None  # Inválido se for 'AAA', 'XXX', 'BBB', etc.

        # 2. Se for padrão BR antigo, verifica também se os 4 números são iguais
        if match_br:
            numeros = placa_normalizada[3:7]
            if len(set(numeros)) == 1:
                return None  # Inválido se for '0000', '1111', etc.
        
        # Se passou em todas as validações, a placa é válida
        return placa_normalizada
    
    else:
        # Se o formato inicial já for inválido
        return None
    
#EMPRESAS = {
#    "001": {
#        "companyName": "IDEAL AUTO PEÇAS LTDA MATRIZ",
#        "companyAddress": "AVENIDA ITAMBÉ, 300 - PATAGONIA",
#        "empresaCnpj": "12.275.282/0001-19"
#    },
#    "002": {
#        "companyName": "IDEAL AUTO PEÇAS LTDA FILIAL",
#        "companyAddress": "AVENIDA PRES. DUTRA, 1070 - JUREMA",
#        "empresaCnpj": "12.275.282/0003-80"
#    },
#    "003": {
#        "companyName": "IDEAL REPARADORA PECAS E SERVIÇOS LTDA",
#        "companyAddress": "RUA TUPINAMBAS, 08 - PATAGONIA",
#        "empresaCnpj": "52.787.630/0001-51"
#    }
#}
def atualizar_empresas():
    """Busca o mapa de empresas via API e atualiza a variável global."""
    global EMPRESAS # Avisa que vamos modificar a variável global
    
    logging.info("Atualizando lista de empresas...")
    resultado = call_app_script("obter_empresas") # Chama a nova ação da API
    
    if resultado.get("status") == "success":
        dados = resultado.get("data", {})
        if isinstance(dados, dict):
            EMPRESAS = dados
            logging.info(f"Lista de empresas atualizada com sucesso. {len(EMPRESAS)} empresas carregadas.")
        else:
            logging.warning("Falha ao atualizar empresas: os dados recebidos não são um dicionário.")
    else:
        logging.error(f"Não foi possível atualizar a lista de empresas: {resultado.get('message', 'Erro desconhecido')}")


async def ajuda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Envia uma mensagem de ajuda completa e adaptada ao tipo de usuário."""
    log_interaction(update)
    user_name = update.effective_user.first_name
    user_id = str(update.effective_user.id) # Pega o ID do usuário para verificar se é admin
    

    # --- Mensagem Base para Todos os Usuários ---
    texto_ajuda = f"Claro, {user_name}! Aqui está um resumo de tudo que eu posso fazer:\n\n"
    
    texto_ajuda += "<b>💬 Comandos por Texto Livre:</b>\n"
    #texto_ajuda += "<code> • novo pedido</code> - Inicia a criação de um novo pedido.\n"
    #texto_ajuda += "<code> • rascunho</code> - Vê e continua um pedido salvo.\n"
    #texto_ajuda += "<code> • pedidos do fornecedor [nome]</code> - Procura por pedidos do Fornecedor.\n"
    #texto_ajuda += "<code> • pedidos da placa [placa]</code> - Procura por pedidos pela Placa.\n"
    #texto_ajuda += "<code> • calcule o imposto de [valor] para [estado de origem]</code> - Calcula rapidamente o valor do imposto ST.\n\n"

    texto_ajuda += "<b>⚙️ Comandos Diretos:</b>\n"
    #texto_ajuda += "<code> • /fornecedores</code> - Mostra a lista de fornecedores ativos.\n"
    texto_ajuda += "<code> • /buscar [Nº Pedido]_[ID Empresa]</code> - Busca por um pedido pelo Numero_ID da Empresa.\n"
    #texto_ajuda += "<code> • /calculoimposto [Valor]_[Estado]</code> - Calcula rapidamente o valor do imposto ST.\n"
    #texto_ajuda += "<code> • /cancelar</code> - Cancela uma operação em andamento (como criar um pedido).\n"
    #texto_ajuda += "<code> • /novofornecedor</code> - Solicita a criação de um novo Fornecedor.\n"

    # --- Seção Especial Apenas para Administradores ---
    if user_id in ADMIN_CHAT_IDS:
        texto_ajuda += "\n"
        texto_ajuda += "👮‍♂️ <b>Comandos de Administração:</b>\n"
        #texto_ajuda += "<code> • /relatorio_pdf [Empresa] [Início] [Fim] [Tipo]</code> - Baixe Relatórios em PDF\n"
        #texto_ajuda += "<code> • /relatorio_xls [Empresa] [Início] [Fim] [Tipo]</code> - Baixe Relatórios em XLSX\n"
        # Adicione aqui outros comandos de admin no futuro
        
    await update.message.reply_text(texto_ajuda, parse_mode="HTML")

ADMIN_CHAT_IDS = set()

def atualizar_admins():
    resultado = call_app_script("obter_admins")
    if resultado.get("status") == "success":
        global ADMIN_CHAT_IDS
        # Extrai apenas os chatIds
        ADMIN_CHAT_IDS = set(str(admin["chatId"]) for admin in resultado.get("data", []))
        logging.info(f"Lista de admins atualizada: {ADMIN_CHAT_IDS}")
    else:
        msg_erro = resultado.get('message', 'Erro desconhecido')
        logging.warning(f"Não foi possível atualizar a lista de admins: {msg_erro}") 

# NOVA FUNÇÃO AJUDANTE
async def _enviar_lista_pedidos(message_able_object, nome_fornecedor: str, pedidos: list):
    """
    Recebe uma lista de pedidos e envia em uma única mensagem
    com uma lista de botões.
    """
   
    # Verifica se não há pedidos
    if not pedidos:
        # Para CallbackQuery, tenta edit primeiro, senão reply
        if hasattr(message_able_object, 'edit_message_text'):
            await message_able_object.edit_message_text(
                f"Nenhum pedido encontrado para <b>{nome_fornecedor}</b>.", 
                parse_mode="HTML"
            )
        else:
            await message_able_object.reply_text(
                f"Nenhum pedido encontrado para <b>{nome_fornecedor}</b>.", 
                parse_mode="HTML"
            )
        return
    
    # Verifica se pedidos é uma lista
    if not isinstance(pedidos, list):
        logging.error(f"[DEBUG] ERRO: pedidos não é uma lista, é {type(pedidos)}")
        await message_able_object.reply_text(
            f"Erro no formato dos dados. Tipo recebido: {type(pedidos)}"
        )
        return
    
    # Cria a lista de botões, um para cada pedido
    botoes = []
    logging.info(f"[DEBUG] Processando {len(pedidos)} pedidos...")
    
    for i, p in enumerate(pedidos):
        logging.info(f"[DEBUG] Pedido {i}: {p}")
        
        if not isinstance(p, dict):
            logging.warning(f"[DEBUG] Pedido {i} não é um dicionário: {type(p)}")
            continue
            
        numero = p.get('numero', 'N/A')
        status = p.get('status', 'N/A')
        empresa_id = p.get('empresaId', '')
        
        logging.info(f"[DEBUG] - numero: {numero}, status: {status}, empresa_id: {empresa_id}")
        
        # Valida se temos dados mínimos para criar o botão
        if numero != 'N/A' and empresa_id:
            botoes.append([
                InlineKeyboardButton(
                    f"🔍 {numero} - {status}",
                    callback_data=f"detalhes:{numero}:{empresa_id}"
                )
            ])
            logging.info(f"[DEBUG] Botão criado para pedido {numero}")
        else:
            logging.warning(f"[DEBUG] Dados insuficientes para criar botão: numero={numero}, empresa_id={empresa_id}")
    
    logging.info(f"[DEBUG] Total de botões criados: {len(botoes)}")
    
    # Se não conseguiu criar nenhum botão válido
    if not botoes:
        texto_erro = f"Encontrei pedidos para <b>{nome_fornecedor}</b>, mas houve um problema com os dados. Tente novamente."
        logging.error("[DEBUG] Nenhum botão foi criado!")
        if hasattr(message_able_object, 'edit_message_text'):
            await message_able_object.edit_message_text(texto_erro, parse_mode="HTML")
        else:
            await message_able_object.reply_text(texto_erro, parse_mode="HTML")
        return

    # Monta a mensagem final
    texto_final = f"📦 Pedidos de <b>{nome_fornecedor}</b> encontrados:"
    markup = InlineKeyboardMarkup(botoes)
    
    # Envia ou edita a mensagem
    try:
        logging.info("[DEBUG] Tentando enviar/editar mensagem com botões...")
        if hasattr(message_able_object, 'edit_message_text'):
            await message_able_object.edit_message_text(
                texto_final,
                reply_markup=markup,
                parse_mode="HTML"
            )
        else:
            await message_able_object.reply_text(
                texto_final,
                reply_markup=markup,
                parse_mode="HTML"
            )
        logging.info("[DEBUG] Mensagem enviada com sucesso!")
    except Exception as e:
        logging.error(f"[DEBUG] ERRO ao enviar mensagem: {e}")
        # Fallback: tenta enviar sem botões
        texto_erro = f"Encontrei {len(pedidos)} pedidos para <b>{nome_fornecedor}</b>, mas houve um erro ao exibir os botões."
        if hasattr(message_able_object, 'reply_text'):
            await message_able_object.reply_text(texto_erro, parse_mode="HTML")

def obter_saudacao_por_horario():
    hora_atual = datetime.now().hour
    if 5 <= hora_atual < 12:
        return "Bom dia"
    elif 12 <= hora_atual < 18:
        return "Boa tarde"
    else:
        return "Boa noite"

async def saudacao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Responde a saudações informais do usuário de forma amigável e dinâmica.
    """
    log_interaction(update)
    user_name = update.effective_user.first_name
    saudacao_horario = obter_saudacao_por_horario()

    opcoes_de_resposta = [
        f"Como posso te ajudar?",
        f"Em que posso ser útil agora?",
        f"Estou a postos! O que vamos verificar hoje?",
        f"O que você precisa?"
    ]

    # Monta uma resposta completa e escolhe um final aleatório
    mensagem = f"{saudacao_horario}, {user_name}! {random.choice(opcoes_de_resposta)}"
    
    await update.message.reply_text(mensagem)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Envia uma mensagem de boas-vindas."""
    log_interaction(update)
    user_name = update.effective_user.first_name
    # 1. Captura todos os dados do usuario
    try:
        # 1. Captura os dados do usuário
        chat_id = update.effective_user.id
        first_name = update.effective_user.first_name
        username = update.effective_user.username or "N/A" # Caso o usuário não tenha username
        data_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # 2. Formata a linha que será salva no arquivo
        linha_para_salvar = f"ID: {chat_id}, Nome: {first_name}, Username: @{username}, Data: {data_hora}\n"

        # 3. Abre o arquivo 'novos_usuarios.txt' em modo 'append' (adicionar) e salva a linha
        # O encoding='utf-8' é importante para nomes com acentos.
        base_dir = os.path.dirname(os.path.abspath(__file__))
        path_usuarios = os.path.join(base_dir, "novos_usuarios.txt")
        with open(path_usuarios, "a", encoding="utf-8") as arquivo:
            arquivo.write(linha_para_salvar)
        
        logging.info(f"Usuário {chat_id} salvo com sucesso no arquivo novos_usuarios.txt")

    except Exception as e:
        logging.error(f"Falha ao salvar usuário {chat_id} no arquivo de texto: {e}")

    saudacao = obter_saudacao_por_horario()
    opcoes_de_saudacao = [
        f"Olá, {user_name}! Pronto para organizar alguns pedidos?",
        f"E aí, {user_name}! O que vamos fazer hoje?",
        f"Bem-vindo(a) de volta, {user_name}! Seu assistente de pedidos está a postos.",
        f"Oi, {user_name}! Como posso te ajudar agora?"
    ]
    
    # Escolhe uma das saudações aleatoriamente
    #mensagem = random.choice(opcoes_de_saudacao)

    mensagem = f"{saudacao}, {user_name}! {random.choice(opcoes_de_saudacao)}\n\nPara começar, que tal tentar `novo pedido` ou ver os comandos em /ajuda?"
    await update.message.reply_text(mensagem)

# --- FUNÇÃO "INTELIGENTE" PARA PROCESSAR MENSAGENS ---
async def processar_mensagem_geral(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trata mensagens de texto, agora com capacidade de entender
    pedidos de cálculo de imposto de forma natural."""
    log_interaction(update)
    if context.user_data.get("aguardando_motivo"):
        return
    
    texto_original = update.message.text.strip()
    # Prepara o ID do usuário UMA VEZ para a verificação de permissão
    user_id = str(update.effective_user.id)
    info_do_usuario = {
        "id": update.effective_user.id,
        "first_name": update.effective_user.first_name,
        "username": update.effective_user.username
    }

    padrao_imposto = re.compile(
         r".*?(\d[\d.,]*).*?\b([A-Z]{2})\b.*", re.IGNORECASE | re.DOTALL)
     #   re.IGNORECASE | re.DOTALL
   # )
    match = padrao_imposto.search(texto_original)

    if 'imposto' in texto_original.lower() and match:
        try:
            valor_str = match.group(1).replace('.', '').replace(',', '.')
            estado_str = match.group(2).upper()
            valor = float(valor_str)
            await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
            # Reutiliza a lógica do comando /calculoimposto
            # O ideal é ter uma função auxiliar para não duplicar código, mas isso funciona.
            await update.message.reply_text(f"Entendi! Calculando imposto para R$ {valor:.2f} vindo de {estado_str}...")
            
            resultado = call_app_script('calcular_imposto_simples', {'valor': valor, 'estado': estado_str})
            
            if resultado.get("status") == "success":
                valor_calculado = resultado.get("valorCalculado", 0)
                aliquota_usada = resultado.get("aliquotaUsada", 0) * 100
                
                mensagem = (
                    f"📊 <b>Resultado do Cálculo de Imposto</b>\n\n"
                    f"<b>Valor Base:</b> R$ {valor:,.2f}\n"
                    f"<b>Estado:</b> {estado_str}\n"
                    f"<b>Alíquota Aplicada:</b> {aliquota_usada:.2f}%\n\n"
                    f"<b>Valor do Imposto:</b> <code>R$ {valor_calculado:,.2f}</code>"
                )
                await update.message.reply_text(mensagem, parse_mode="HTML")
            else:
                opcoes_resposta = [
                    "Opa, não entendi o que você quis dizer. Que tal tentar `novo pedido` ou usar o comando /ajuda?",
                    "Hmm, não captei essa. Lembre-se que você pode ver tudo que eu faço com o comando /ajuda.",
                    "Desculpe, não processei seu pedido. Para ver as opções disponíveis, é só chamar o /ajuda."
                ]
                await update.message.reply_text(random.choice(opcoes_resposta))

            return # Termina a execução aqui
            
        except (ValueError, IndexError):
            await update.message.reply_text("Entendi que você quer calcular um imposto, mas não consegui identificar o valor e o estado. Tente novamente ou use /calculoimposto VALOR_ESTADO.")
            return

    # Se não for um cálculo de imposto, continuamos para as outras verificações
    texto_lower = texto_original.lower()

    # --- 3. OUVINTE PARA "RASCUNHO" ---
    if 'rascunho' in texto_lower:
        # Chama a mesma função do comando /rascunhos
        await rascunhos(update, context)
        return # Termina aqui, pois a função 'rascunhos' já envia a resposta

    elif texto_lower.startswith("pedidos do fornecedor"):
        nome_fornecedor = texto_original[len("pedidos do fornecedor"):].strip()
        if not nome_fornecedor:
            await update.message.reply_text("Claro! Me diga o nome do fornecedor que você deseja buscar.")
            return
        await asyncio.to_thread(atualizar_fornecedores) 
        # Busca por parte do nome, ignorando maiúsculas/minúsculas
        fornecedores_encontrados = [
            f for f in FORNECEDORES if nome_fornecedor.lower() in f.lower()
        ]
        if not fornecedores_encontrados:
            await update.message.reply_text(
                f"🤔 Hmm, não encontrei nenhum fornecedor que contenha '{nome_fornecedor}'.\n\n"
                f"Tente um nome diferente ou use o comando /fornecedores para ver a lista completa."
            )
            return
        # Se houver mais de um, pode pedir para o usuário escolher, mas aqui pega o primeiro
        fornecedor_escolhido = fornecedores_encontrados[0]
        context.user_data['ultima_busca_fornecedor'] = fornecedor_escolhido

        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
        await update.message.reply_text(f"🔎 Ok, buscando pedidos para <b>{fornecedor_escolhido}</b>. Um momento...", parse_mode="HTML")
        resultado = call_app_script('buscar_por_fornecedor', {
            'nomeFornecedor': fornecedor_escolhido,
            'userInfo': info_do_usuario # Envia info do usuário
        })
        
        pedidos_brutos = resultado.get('data', [])

        # --- FILTRO DE SEGURANÇA APLICADO AQUI ---
        if user_id not in ADMIN_CHAT_IDS:
            # Se o usuário NÃO é admin, filtramos a lista
            pedidos_para_mostrar = [
                p for p in pedidos_brutos 
                if p.get('status', '').upper() != 'AGUARDANDO APROVACAO'
            ]
            logging.info(f"Filtro aplicado para usuário {user_id}. Mostrando {len(pedidos_para_mostrar)} de {len(pedidos_brutos)} pedidos.")
        else:
            # Se o usuário É admin, ele vê a lista completa
            pedidos_para_mostrar = pedidos_brutos
            logging.info(f"Admin {user_id} acessando. Mostrando todos os {len(pedidos_brutos)} pedidos.")
        # --- FIM DO FILTRO ---
        logging.info(f"Tipo de pedidos recebido: {type(pedidos_para_mostrar)} - Conteúdo: {pedidos_para_mostrar}")
        await _enviar_lista_pedidos(update.message, fornecedor_escolhido, pedidos_para_mostrar)
        return
    elif texto_original.startswith("pedidos da placa"):
        placa_raw = texto_original[len("pedidos da placa"):].strip()
        if not placa_raw:
            await update.message.reply_text("Certo! Por favor, me informe a placa do veículo.")
            return
        placa_normalizada = normalizar_placa(placa_raw)
        safe_placa = html.escape(placa_raw.upper())
        await update.message.reply_text(f"🔎 Entendido! Deixa eu ver o que encontro para a placa <b>{safe_placa}</b>...", parse_mode="HTML")
        resultado = call_app_script('buscar_por_placa', {
            'placaVeiculo': placa_normalizada,
            'userInfo': info_do_usuario # Envia info do usuário
        })
        logging.info(f"Retorno do App Script (placa): {resultado}")
        pedidos_brutos = resultado.get('data', [])
        if user_id not in ADMIN_CHAT_IDS:
            pedidos_para_mostrar = [p for p in pedidos_brutos if p.get('status', '').upper() != 'AGUARDANDO APROVACAO']
        else:
            pedidos_para_mostrar = pedidos_brutos
        
        await _enviar_lista_pedidos(update.message, placa_raw.upper(), pedidos_para_mostrar)
        return
    else:
        opcoes_resposta = [
            "Opa, não entendi o que você quis dizer.",
            "Hmm, não captei essa.",
            "Desculpe, não processei seu pedido."
        ]
        await update.message.reply_text(f"{random.choice(opcoes_resposta)} Que tal tentar `novo pedido` ou ver todos os comandos em /ajuda?")

# Exemplo de lista fixa, pode ser dinâmica via API
FORNECEDORES = {}

def atualizar_fornecedores():
    """Busca o MAPA de fornecedores via API e o atualiza na variável global."""
    # A palavra-chave 'global' é essencial para modificar a variável de fora da função
    global FORNECEDORES
    
    resultado = call_app_script('criarMapaDeFornecedoresv2')
    mapa = resultado.get('data', {})
    
    # O 'if' e o 'else' estão perfeitamente alinhados
    if isinstance(mapa, dict):
        # As linhas dentro do 'if' estão indentadas com 4 espaços
        FORNECEDORES = mapa
        logging.info(f"Mapa de fornecedores atualizado com {len(FORNECEDORES)} registros.")
    else:
        # As linhas dentro do 'else' estão indentadas com 4 espaços
        FORNECEDORES = {}
        logging.warning("Não foi possível atualizar o mapa de fornecedores.")

async def fornecedores(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)

    if context.bot_data.get('fornecedores_list'):
        logging.info("Carregando fornecedores do cache.")
        fornecedores_keys = context.bot_data['fornecedores_list']
    else:
        # Se não estiver no cache, busca na API (comportamento de fallback)
        logging.info("Cache de fornecedores vazio. Buscando da API...")

    await update.message.reply_text("⏳ Minha equipe está buscando a lista de fornecedores...")
    await asyncio.to_thread(atualizar_fornecedores)
    fornecedores_keys = sorted(list(FORNECEDORES.keys()), key=lambda s: s.lower())
    context.bot_data['fornecedores_list'] = fornecedores_keys

    if not fornecedores_keys:
        await update.message.reply_text("Nenhum fornecedor encontrado.")
        return

    # Guarda a lista completa no bot_data para ser acessada pelas outras páginas
    #context.bot_data['fornecedores_list'] = fornecedores_keys

    # Cria o teclado para a primeira página (page=0)
    keyboard = await build_paginated_keyboard(fornecedores_keys, page=0)

    await update.message.reply_text(
        "Encontrei esta lista, por favor, escolha um fornecedor (Página 1):",
        reply_markup=keyboard
    )

async def buscar_fornecedor_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    query = update.callback_query
    await query.answer()
    try:
        # Agora usa o índice em vez do nome diretamente
        _, indice_str = query.data.split(":", 1)
        indice = int(indice_str)
        
        # Recupera a lista salva
        fornecedores_list = context.bot_data.get('fornecedores_list', [])
        
        if indice >= len(fornecedores_list):
            await query.edit_message_text("Erro: Fornecedor não encontrado. Tente buscar novamente.")
            return
            
        nome_fornecedor = fornecedores_list[indice]
        
        context.user_data['ultima_busca_fornecedor'] = nome_fornecedor
        await query.edit_message_text(f"🔎 Buscando pedidos do fornecedor <b>{nome_fornecedor}</b>...", parse_mode="HTML")
        
        resultado = call_app_script('buscar_por_fornecedor', {'nomeFornecedor': nome_fornecedor})
        pedidos = resultado.get('data', [])
        await _enviar_lista_pedidos(query.message, nome_fornecedor, pedidos)
        
    except (ValueError, IndexError) as e:
        logging.error(f"Erro no callback do fornecedor: {e}")
        await query.edit_message_text("Erro ao processar sua seleção. Tente novamente.")

# Adicione esta nova função de callback
async def fornecedores_page_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Lida com a navegação de páginas da lista de fornecedores."""
    query = update.callback_query
    await query.answer()

    # Extrai o número da página do callback_data (ex: "page_fornecedores:1")
    _, page_str = query.data.split(":", 1)
    page = int(page_str)

    # Recupera a lista completa de fornecedores que salvamos
    fornecedores_list = context.bot_data.get('fornecedores_list', [])
    if not fornecedores_list:
        await query.edit_message_text("Erro: a lista de fornecedores expirou. Por favor, use /fornecedores novamente.")
        return

    # Cria o novo teclado para a página solicitada
    keyboard = await build_paginated_keyboard(fornecedores_list, page=page)

    # Edita a mensagem original com a nova página e o novo teclado
    await query.edit_message_text(
        f"Escolha um fornecedor (Página {page + 1}):",
        reply_markup=keyboard
    )

CAD_CNPJ, CONFIRMAR_CADASTRO, SELECIONAR_CONDICAO, SELECIONAR_FORMA = range(20, 24)

async def novo_fornecedor_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Inicia a conversa para cadastrar um novo fornecedor via CNPJ (apenas admins)."""
    user_id = str(update.effective_user.id)
    if user_id not in ADMIN_CHAT_IDS:
        await update.message.reply_text("Desculpe, este comando é apenas para administradores.")
        return ConversationHandler.END

    await update.message.reply_text("Ok, vamos cadastrar um novo fornecedor. Por favor, digite o CNPJ.")
    return CAD_CNPJ

async def receber_cnpj(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Recebe o CNPJ, consulta na API e pede a confirmação do usuário."""
    #cnpj = update.message.text.strip()
    #await update.message.reply_text(f"🔎 Buscando dados para o CNPJ {cnpj}")
    
    cnpj = re.sub(r'\D', '', update.message.text)  # remove tudo que não for número
    if len(cnpj) != 14:
        await update.message.reply_text("❌ CNPJ inválido! Certifique-se de digitar 14 dígitos.")
        return CAD_CNPJ

    await update.message.reply_text(f"🔎 Buscando dados para o CNPJ {cnpj}")
    logging.info(f"[DEBUG] Chamando App Script com CNPJ: {cnpj} e ação: _api_consultarCnpjEopcoes")
    resultado = call_app_script("consultar_cnpj_e_opcoes", {"cnpj": cnpj})
    logging.info(f"[DEBUG] Resultado do App Script: {resultado}")

    if resultado.get("status") == "success":
        dados_completos = resultado.get("data", {})
        logging.info(f"Dados recebidos do App Script: {dados_completos}")
        context.user_data['dados_fornecedor_confirmar'] = dados_completos.get('dadosFornecedor')
        context.user_data['opcoes_condicoes'] = dados_completos.get('opcoesCondicoes')
        context.user_data['opcoes_formas'] = dados_completos.get('opcoesFormas')
        
        dados_fornecedor = context.user_data['dados_fornecedor_confirmar']
        mensagem_confirmacao = (
            f"<b>Dados Encontrados:</b>\n\n"
            f"<b>Razão Social:</b> {dados_fornecedor.get('razaoSocial', 'N/A')}\n"
            f"<b>Nome Fantasia:</b> {dados_fornecedor.get('nomeFantasia', 'N/A')}\n"
            f"<b>Endereço:</b> {dados_fornecedor.get('endereco', 'N/A')}\n"
            f"<b>Cidade/UF:</b> {dados_fornecedor.get('cidade', '')}/{dados_fornecedor.get('uf', '')}\n\n"
            f"Posso prosseguir com o cadastro deste fornecedor?"
        )
        botoes = [[
            InlineKeyboardButton("✅ Sim", callback_data="confirmar_cadastro_sim"),
            InlineKeyboardButton("❌ Não", callback_data="confirmar_cadastro_nao")
        ]]
        await update.message.reply_text(mensagem_confirmacao, reply_markup=InlineKeyboardMarkup(botoes), parse_mode="HTML")
        return CONFIRMAR_CADASTRO
    else:
        await update.message.reply_text(f"❌ Erro ao consultar CNPJ: {resultado.get('message', 'Falha.')}")
        return ConversationHandler.END
    
async def confirmar_cadastro_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Lida com a confirmação e, se for 'Sim', pede a Condição de Pagamento."""
    query = update.callback_query
    await query.answer()

    if query.data == 'confirmar_cadastro_nao':
        await query.edit_message_text("Cadastro cancelado.")
        return ConversationHandler.END
        
    condicoes = context.user_data.get('opcoes_condicoes', [])
    if not condicoes:
        await query.edit_message_text("❌ Nenhuma condição de pagamento encontrada. Cadastro cancelado.")
        return ConversationHandler.END

    botoes = [[InlineKeyboardButton(c, callback_data=f"cad_condicao:{c}")] for c in condicoes]
    await query.edit_message_text(
        "✅ Dados confirmados. Agora, selecione a Condição de Pagamento padrão:",
        reply_markup=InlineKeyboardMarkup(botoes)
    )
    return SELECIONAR_CONDICAO

async def selecionar_condicao_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Recebe a Condição e pede a Forma de Pagamento."""
    query = update.callback_query
    await query.answer()
    
    condicao_escolhida = query.data.split(":", 1)[1]
    context.user_data['dados_fornecedor_confirmar']['condicaoPagamento'] = condicao_escolhida
    
    formas = context.user_data.get('opcoes_formas', [])
    if not formas:
        await query.edit_message_text("❌ Nenhuma forma de pagamento encontrada. Cadastro cancelado.")
        return ConversationHandler.END

    botoes = [[InlineKeyboardButton(f, callback_data=f"cad_forma:{f}")] for f in formas]
    await query.edit_message_text(
        f"Condição: '{condicao_escolhida}'.\n\nAgora, selecione a Forma de Pagamento:",
        reply_markup=InlineKeyboardMarkup(botoes)
    )
    return SELECIONAR_FORMA

async def selecionar_forma_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Recebe a Forma, envia tudo para a API e encerra."""
    query = update.callback_query
    await query.answer()

    forma_escolhida = query.data.split(":", 1)[1]
    context.user_data['dados_fornecedor_confirmar']['formaPagamento'] = forma_escolhida
    
    await query.edit_message_text("Ótimo! Enviando dados finais para o sistema...")
    
    resultado = call_app_script("finalizar_cadastro_fornecedor", {
        "fornecedorData": context.user_data['dados_fornecedor_confirmar']
    })

    await query.edit_message_text(resultado.get("message", "Processo finalizado."))
    context.user_data.clear()
    return ConversationHandler.END

async def cad_fornecedor_cnpj(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Recebe o CNPJ, busca as opções de pagamento e pede a Condição."""
    context.user_data['novo_fornecedor'] = {'cnpj': update.message.text.strip()}
    
    await update.message.reply_text("Buscando opções de pagamento...")
    
    resultado = call_app_script("obter_opcoes_pagamento")
    if resultado.get("status") != "success":
        await update.message.reply_text(f"❌ Erro ao buscar opções: {resultado.get('message')}. Ação cancelada.")
        return ConversationHandler.END
        
    opcoes = resultado.get('data', {})
    context.user_data['opcoes_pagamento'] = opcoes

    condicoes = opcoes.get('condicoes', [])
    if not condicoes:
        await update.message.reply_text("❌ Nenhuma condição de pagamento encontrada na planilha 'Config'. Ação cancelada.")
        return ConversationHandler.END

    botoes = [[InlineKeyboardButton(c, callback_data=f"cad_condicao:{c}")] for c in condicoes]
    await update.message.reply_text(
        "Por favor, selecione a Condição de Pagamento padrão para este fornecedor:",
        reply_markup=InlineKeyboardMarkup(botoes)
    )
    return CAD_FORNECEDOR_CONDICAO

async def CAD_FORNECEDOR_CONDICAO(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Recebe a Condição, e pede a Forma de Pagamento."""
    query = update.callback_query
    await query.answer()
    
    condicao_escolhida = query.data.split(":", 1)[1]
    context.user_data['novo_fornecedor']['condicaoPagamento'] = condicao_escolhida
    
    formas = context.user_data.get('opcoes_pagamento', {}).get('formas', [])
    if not formas:
        await query.edit_message_text("❌ Nenhuma forma de pagamento encontrada na planilha 'Config'. Ação cancelada.")
        return ConversationHandler.END

    botoes = [[InlineKeyboardButton(f, callback_data=f"cad_forma:{f}")] for f in formas]
    await query.edit_message_text(
        f"Condição: '{condicao_escolhida}'.\n\nAgora, selecione a Forma de Pagamento:",
        reply_markup=InlineKeyboardMarkup(botoes)
    )
    return CAD_FORNECEDOR_FORMA

async def CAD_FORNECEDOR_FORMA(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Recebe a Forma, envia tudo para a API e encerra."""
    query = update.callback_query
    await query.answer()

    forma_escolhida = query.data.split(":", 1)[1]
    context.user_data['novo_fornecedor']['formaPagamento'] = forma_escolhida
    
    await query.edit_message_text("✅ Ótimo! Tenho todos os dados. Estou verificando e cadastrando...")
    
    resultado = call_app_script("cadastrar_fornecedor_via_cnpj", context.user_data['novo_fornecedor'])

    if resultado.get("status") == "success":
        await query.edit_message_text(resultado.get("message"))
    else:
        await query.edit_message_text(f"❌ Erro: {resultado.get('message')}")
        
    context.user_data.clear()
    return ConversationHandler.END

# --- HANDLERS DE COMANDOS E CALLBACKS ---
async def buscar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    """Handler para o comando /buscar NUMERO_EMPRESA."""
    try:
        if not context.args or '_' not in context.args[0]:
            raise ValueError("Formato inválido")
        args = context.args[0]
        numero_pedido, empresa_id = args.split('_')
        
        await update.message.reply_text(f"🔎 Buscando pedido <b>{numero_pedido}</b>...", parse_mode='HTML')
        # Prepara as informações do usuário para enviar ao backend
        info_do_usuario = {
            "id": update.effective_user.id,
            "first_name": update.effective_user.first_name,
            "username": update.effective_user.username
        }

        # Adiciona o userInfo aos parâmetros da chamada
        params_busca = {
            "mainSearch": numero_pedido, 
            "empresaId": empresa_id,
            "userInfo": info_do_usuario
        }
        print("--> Enviando para o Apps Script:", params_busca)
        resultado = call_app_script('buscar_pedido', params_busca)
        print("<-- Resultado do Apps Script:", resultado)
        if resultado.get("status") == "timeout":
            await update.message.reply_text("⏳ O servidor demorou demais para responder. Tente novamente mais tarde.")
            return
        if resultado.get("status") == "network_error":
            await update.message.reply_text("❌ Falha de comunicação com o sistema. Tente novamente mais tarde.")
            return
        
        if resultado.get("status") == "success":
            botoes = [[
                InlineKeyboardButton("🔎 Ver detalhes", callback_data=f"detalhes:{numero_pedido}:{empresa_id}"),
                #InlineKeyboardButton("✅ Aprovar", callback_data=f"aprovar:{numero_pedido}:{empresa_id}"),
                #InlineKeyboardButton("❌ Rejeitar", callback_data=f"rejeitar:{numero_pedido}:{empresa_id}"),
                InlineKeyboardButton("📄 PDF", callback_data=f"pdf:{numero_pedido}:{empresa_id}")
            ]]
        
            await update.message.reply_text(resultado['data'], reply_markup=InlineKeyboardMarkup(botoes), parse_mode='HTML')
        else:
            await update.message.reply_text(resultado.get('data', 'Ocorreu um erro.'), parse_mode='HTML')
    except (IndexError, ValueError):
            await update.message.reply_text("😕 Formato inválido. Use: `/buscar NUMEROPEDIDO_IDEMPRESA`")

async def relatorio_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logging.info("Handler relatorio_pdf chamado")
    log_interaction(update)
    """
    Gera relatório PDF de compras.
    Uso: /relatorio_pdf EMPRESA_ID DATA_INICIO DATA_FIM TIPO [filtro=valor ...]
    Aceita datas em DD/MM/AAAA ou AAAA-MM-DD e tipos em PT-BR (detalhado/financeiro).
    """
    def _parse_date(date_str):
        # Tenta dd/mm/YYYY então ISO YYYY-MM-DD
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        raise ValueError("Formato inválido")

    def _map_report_type(tipo_raw):
        mapping = {
            "detalhado": "detailed", "detalhe": "detailed", "detailed": "detailed", "d": "detailed",
            "financeiro": "financial", "financial": "financial", "f": "financial"
        }
        key = tipo_raw.strip().lower()
        return mapping.get(key, "detailed")  # se não reconhecer, envia como detalhado

    user_id = str(update.effective_user.id)
    if user_id not in ADMIN_CHAT_IDS:
        await update.message.reply_text("❌ Eu não tenho permissão para gerar esses relatórios para você.")
        return

    try:
        args = context.args or []
        if len(args) < 4:
            await update.message.reply_text(
                "Use: /relatorio_pdf EMPRESA_ID DATA_INICIO DATA_FIM TIPO\n"
                "Exemplo (pt-br): /relatorio_pdf 001 01/08/2025 12/08/2025 detalhado\n"
                "Exemplo (pt-br): /relatorio_pdf 001 01/08/2025 12/08/2025 detalhado fornecedor=ACME\n"
                "Ou (iso): /relatorio_pdf 001 2025-08-01 2025-08-12 financial"
            )
            return
        empresa_id, data_inicio_raw, data_fim_raw, tipo_raw = args[:4]
    except Exception as e:
        logging.exception("Erro ao ler argumentos de /relatorio_pdf")
        await update.message.reply_text("Erro ao processar os argumentos do comando. Verifique a sintaxe.")
        return

    empresa_info = EMPRESAS.get(empresa_id)
    if not empresa_info:
        await update.message.reply_text(f"Empresa '{empresa_id}' não encontrada.")
        return

    try:
        data_inicio = _parse_date(data_inicio_raw)
        data_fim = _parse_date(data_fim_raw)
    except ValueError:
        await update.message.reply_text("Formato de data inválido. Use DD/MM/AAAA ou AAAA-MM-DD (ex: 01/08/2025).")
        return

    report_type = _map_report_type(tipo_raw)

    # parse filtros opcionais chave=valor (normaliza chaves para lowercase)
    filtros_opcionais = {}
    for extra in args[4:]:
        if "=" in extra:
            k, v = extra.split("=", 1)
            filtros_opcionais[k.strip().lower()] = v.strip()

    # Aceita chaves em PT-BR ou EN: 'fornecedor' / 'supplier', e 'status'.
    supplier = filtros_opcionais.pop('fornecedor', None) or filtros_opcionais.pop('supplier', None)
    status = filtros_opcionais.pop('status', None)

    params = {
        "companyId": empresa_id,
        "companyName": empresa_info.get("empresa", ""),
        "companyAddress": empresa_info.get("endereco", ""),
        "empresaCnpj": empresa_info.get("cnpj", ""),
        "startDate": data_inicio,   # sempre enviamos YYYY-MM-DD para o backend
        "endDate": data_fim,
        "reportType": report_type,  # já mapeado de PT-BR para o valor do backend
        "supplier": supplier,       # pode ser None (significa "todos")
        "status": status,
        "filters": filtros_opcionais
    }

    await update.message.reply_text("⏳ Estou trabalhando, gerando seu relatório PDF, aguarde...")
    resultado = call_app_script('relatorio_pdf', params)
    logging.info(f"Retorno do App Script (relatorio_pdf): {resultado}")
    if isinstance(resultado, dict) and resultado.get("status") == "success" and resultado.get("url"):
        await update.message.reply_text(
            f"Seu relatório foi gerado. Baixe aqui:\n[PDF]({resultado['url']})",
            parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(
            f"Erro ao gerar relatório: {resultado.get('message', 'Erro desconhecido')}"
        )

async def relatorio_xls(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Gera relatório XLSX de compras.
    Uso: /relatorio_xls EMPRESA_ID DATA_INICIO DATA_FIM TIPO [filtro=valor ...]
    Aceita datas em DD/MM/AAAA ou AAAA-MM-DD e tipos em PT-BR (detalhado/financeiro).
    """
    log_interaction(update)
    def _parse_date(date_str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        raise ValueError("Formato inválido")

    def _map_report_type(tipo_raw):
        mapping = {
            "detalhado": "detailed", "detalhe": "detailed", "detailed": "detailed", "d": "detailed",
            "financeiro": "financial", "financial": "financial", "f": "financial"
        }
        key = tipo_raw.strip().lower()
        return mapping.get(key, "detailed")

    user_id = str(update.effective_user.id)
    if user_id not in ADMIN_CHAT_IDS:
        await update.message.reply_text("❌ Eu não tenho permissão para gerar esses relatórios para você.")
        return

    try:
        args = context.args or []
        if len(args) < 4:
            await update.message.reply_text(
                "Use: /relatorio_xls EMPRESA_ID DATA_INICIO DATA_FIM TIPO\n"
                "Exemplo (pt-br): /relatorio_xls 001 01/08/2025 12/08/2025 detalhado\n"
                "Exemplo (pt-br): /relatorio_pdf 001 01/08/2025 12/08/2025 detalhado fornecedor=ACME\n"
                "Ou (iso): /relatorio_xls 001 2025-08-01 2025-08-12 financial"
            )
            return
        empresa_id, data_inicio_raw, data_fim_raw, tipo_raw = args[:4]
    except Exception:
        logging.exception("Erro ao ler argumentos de /relatorio_xls")
        await update.message.reply_text("Erro ao processar os argumentos do comando. Verifique a sintaxe.")
        return

    empresa_info = EMPRESAS.get(empresa_id)
    if not empresa_info:
        await update.message.reply_text(f"Empresa '{empresa_id}' não encontrada.")
        return

    try:
        data_inicio = _parse_date(data_inicio_raw)
        data_fim = _parse_date(data_fim_raw)
    except ValueError:
        await update.message.reply_text("Formato de data inválido. Use DD/MM/AAAA ou AAAA-MM-DD (ex: 01/08/2025).")
        return

    report_type = _map_report_type(tipo_raw)

    filtros_opcionais = {}
    for extra in args[4:]:
        if "=" in extra:
            k, v = extra.split("=", 1)
            filtros_opcionais[k.strip().lower()] = v.strip()

    supplier = filtros_opcionais.pop('fornecedor', None) or filtros_opcionais.pop('supplier', None)
    status = filtros_opcionais.pop('status', None)

    params = {
        "companyId": empresa_id,
        "companyName": empresa_info.get("empresa", ""),
        "companyAddress": empresa_info.get("endereco", ""),
        "empresaCnpj": empresa_info.get("cnpj", ""),
        "startDate": data_inicio,
        "endDate": data_fim,
        "reportType": report_type,
        "supplier": supplier,
        "status": status,
        "filters": filtros_opcionais
    }

    await update.message.reply_text("⏳ Estou trabalhando, gerando seu relatório XLSX, aguarde...")
    try:
        resultado = call_app_script('relatorio_xls', params)
        if isinstance(resultado, dict) and resultado.get("status") == "success" and resultado.get("url"):
            await update.message.reply_text(
                f"Seu relatório foi gerado. Baixe aqui:\n[XLSX]({resultado['url']})",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"Erro ao gerar relatório: {resultado.get('message', 'Erro desconhecido')}")
    except Exception as e:
        logging.exception("Erro ao chamar App Script para relatorio_xls")
        await update.message.reply_text(f"Erro ao gerar relatório XLSX: {e}")

async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    """Manipula cliques em botões (detalhes, aprovar, voltar)."""
    query = update.callback_query
    await query.answer()

    if query.data.startswith('voltar_para_lista'):
        partes = query.data.split(':', 1)
        nome_fornecedor = partes[1] if len(partes) > 1 else context.user_data.get('ultima_busca_fornecedor')
        if not nome_fornecedor:
            await query.edit_message_text("Não consigo me lembrar da sua última busca. Por favor, pesquise novamente.")
            return
        await query.edit_message_text(
            f"🔙 Voltando para a lista de pedidos do fornecedor <b>{nome_fornecedor}</b>.",
            parse_mode="HTML"
        )
        resultado = call_app_script('buscar_por_fornecedor', {'nomeFornecedor': nome_fornecedor})
        pedidos = resultado.get('data', [])
        logging.info(f"Tipo de pedidos recebido: {type(pedidos)} - Conteúdo: {pedidos}")
        await _enviar_lista_pedidos(query.message, nome_fornecedor, pedidos)
        return

    action, numero_pedido, empresa_id = query.data.split(':')
    admin_info = {"first_name": query.from_user.first_name, "id": query.from_user.id}
    params = {"numeroPedido": numero_pedido, "empresaId": empresa_id, "adminInfo": admin_info}
    params["nomeFornecedor"] = context.user_data.get('ultima_busca_fornecedor', '')

    if action == 'detalhes':
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
        await query.edit_message_text(f"Estou trabalhando, buscando detalhes do pedido <b>{numero_pedido}</b>...", parse_mode='HTML')
        resultado = call_app_script("obter_detalhes_pedido", params)
        if resultado.get("status") == "timeout":
            await query.edit_message_text("⏳ O servidor demorou demais para responder. Tente novamente mais tarde.")
            return
        if resultado.get("status") == "network_error":
            await query.edit_message_text("❌ Falha de comunicação com o sistema. Tente novamente mais tarde.")
            return
        botoes = [[InlineKeyboardButton("⬅️ Voltar para a lista", callback_data=f"voltar_para_lista:{params.get('nomeFornecedor', '')}")]]
        await query.edit_message_text(
            resultado["data"], reply_markup=InlineKeyboardMarkup(botoes), parse_mode="HTML"
        )       
    elif action == 'aprovar':
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
        await query.edit_message_text(f"Processando aprovação para o pedido {numero_pedido}...")
        resultado = call_app_script("aprovar_pedido", params)
        await query.edit_message_text(resultado["data"], parse_mode="HTML")
    
    if action == 'pdf':
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
        await query.edit_message_text("Estou gerando PDF do pedido, aguarde...")
        resultado = call_app_script("gerar_pdf_pedido", {"numeroPedido": numero_pedido, "empresaId": empresa_id})

        # TRATAMENTO DE TIMEOUT E ERRO DE REDE
        if resultado.get("status") == "timeout":
            await query.edit_message_text("⏳ O servidor demorou demais para responder. Tente novamente mais tarde.")
            return
        if resultado.get("status") == "network_error":
            await query.edit_message_text("❌ Falha de comunicação com o sistema. Tente novamente mais tarde.")
            return
    
        pdf_url = resultado.get("pdfUrl")
        logging.info(f"URL do PDF recebido: {pdf_url}")
        if pdf_url:
            response = requests.get(pdf_url)
            if response.status_code == 200:
                pdf_bytes = io.BytesIO(response.content)
                pdf_bytes.name = f"Pedido_{numero_pedido}.pdf"
                await context.bot.send_document(
                    chat_id=query.message.chat_id,
                    document=pdf_bytes,
                    caption=f"📄 PDF do pedido {numero_pedido}"
                )
                await query.edit_message_text("PDF enviado com sucesso!")
            else:
                await query.edit_message_text("Falha ao baixar o PDF do servidor.")
        else:
            await query.edit_message_text(f"Falha ao gerar PDF: {resultado.get('message', 'Erro desconhecido')}")

# --- HANDLERS DA CONVERSA DE REJEIÇÃO ---

async def rejeitar_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Inicia o processo de rejeição."""
    log_interaction(update)
    query = update.callback_query
    await query.answer()
    
    _, numero_pedido, _ = query.data.split(':')
    context.user_data['rejeitando_pedido'] = numero_pedido
    context.user_data['aguardando_motivo'] = True  # Ativa a flag

    await query.edit_message_text(
        f"❌ Rejeitando Pedido <b>{numero_pedido}</b>...\n\nPor favor, digite e envie o motivo da rejeição.",
        parse_mode='HTML'
    )
    return MOTIVO_REJEICAO

async def receber_motivo_rejeicao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    motivo = update.message.text
    numero_pedido = context.user_data.get("rejeitando_pedido")  # Corrigido aqui!
    admin_info = {
        "first_name": update.message.from_user.first_name,
        "id": update.message.from_user.id,
    }

    await update.message.reply_text(
        f"⏳ Estou registrando sua rejeição para o pedido {numero_pedido}, um momento!"
    )

    params = {
        "numeroPedido": numero_pedido,
        "motivoRejeicao": motivo,
        "adminInfo": admin_info,
    }
    resultado = call_app_script("rejeitar_pedido", params)
    await update.message.reply_text(resultado["data"], parse_mode="HTML")

     # Envia notificação ao criador do pedido, se o chat_id estiver presente
    criador_chat_id = resultado.get("criador_chat_id") or resultado.get("criador_id")
    logging.info(f"Valor de criador_chat_id recebido: {criador_chat_id}")
    if criador_chat_id:
        try:
            logging.info(f"Tentando enviar mensagem de rejeição para chat_id: {criador_chat_id}")
            await context.bot.send_message(
                chat_id=criador_chat_id,
                text=f"❌ <b>Pedido Rejeitado.</b>\n\nSeu pedido <b>Nº {numero_pedido}</b> foi rejeitado.\n<b>Motivo:</b> <i>{motivo or 'N/A'}</i>",
                parse_mode="HTML"
            )
            logging.info(f"Mensagem enviada com sucesso para {criador_chat_id}")
        except Exception as e:
            logging.error(f"Falha ao enviar notificação ao criador do pedido: {e}")
    else:
        logging.warning("Campo criador_chat_id não encontrado ou vazio na resposta do App Script.")

    context.user_data.pop('aguardando_motivo', None) 
    context.user_data.pop('rejeitando_pedido', None)
    return ConversationHandler.END

async def cancelar_conversa(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    await update.message.reply_text("Ação cancelada.")
    context.user_data.clear()
    return ConversationHandler.END

async def erro_handler(update, context):
    logging.error(f"Erro inesperado: {context.error}")

CRIAR_EMPRESA, CRIAR_FORNECEDOR, CRIAR_ITENS, CRIAR_CONFIRMA = range(4)
# Novos estados para a criação de itens
ITEM_DESCRICAO, ITEM_PROD_FORNECEDOR, ITEM_QTD, ITEM_UNIDADE, ITEM_PRECO, ITEM_NOVO_OU_FIM = range(4, 10)
CRIAR_PLACA, CRIAR_NOME_VEICULO, CRIAR_OBSERVACOES = range(10, 13)
CRIAR_OBSERVACOES_OPCIONAL = range(13, 14) 

async def item_descricao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    context.user_data['item_atual']['descricao'] = update.message.text
    await update.message.reply_text("Qual o <b>código do fornecedor</b> para este item?", parse_mode="HTML")
    return ITEM_PROD_FORNECEDOR

async def item_prod_fornecedor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    context.user_data['item_atual']['produtoFornecedor'] = update.message.text
    nome_item = context.user_data['item_atual']['descricao']
    await update.message.reply_text(f"Ok. Para o item '{nome_item}', qual a <b>quantidade</b>?", parse_mode="HTML")
    return ITEM_QTD

async def item_qtd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    try:
        # Tenta converter vírgula para ponto para aceitar ambos os formatos
        quantidade = float(update.message.text.replace(',', '.'))
        context.user_data['item_atual']['quantidade'] = quantidade
        await update.message.reply_text("Qual a <b>unidade</b> de medida (UN, CX, MT, L, RL)?", parse_mode="HTML")
        return ITEM_UNIDADE
    except ValueError:
        await update.message.reply_text("❌ Quantidade inválida. Por favor, envie apenas números.")
        return ITEM_QTD # Permanece no mesmo estado para o usuário corrigir

async def item_unidade(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    context.user_data['item_atual']['unidade'] = update.message.text.upper()
    nome_item = context.user_data['item_atual']['descricao']
    await update.message.reply_text(f"Certo. E qual o <b>preço unitário</b> de '{nome_item}'?", parse_mode="HTML")
    return ITEM_PRECO

async def item_preco(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    try:
        preco = float(update.message.text.replace(',', '.'))

        grupo_fornecedor = context.user_data.get('grupo_fornecedor', '')
        if grupo_fornecedor == 'OBRIGATORIO' and preco <= 0:
            await update.message.reply_text("❌ Para este fornecedor, o preço do item deve ser maior que zero. Por favor, informe um valor válido.")
            return ITEM_PRECO # Permanece no mesmo estado para correção
        
        context.user_data['item_atual']['precoUnitario'] = preco
        
        # Calcula o total do item
        quantidade = context.user_data['item_atual']['quantidade']
        totalItem = quantidade * preco
        context.user_data['item_atual']['totalItem'] = totalItem
        
        # (Opcional) Aqui você pode adicionar a lógica de cálculo de imposto
        context.user_data['item_atual']['icmsSt'] = 0 # Valor padrão
        
        # Adiciona o item completo à lista de itens do pedido
        context.user_data['itens'].append(context.user_data['item_atual'])
        
        await update.message.reply_text(f"✅ Item '{context.user_data['item_atual']['descricao']}' adicionado!")

        # Pergunta o que fazer a seguir com botões
        botoes = [
            [InlineKeyboardButton("➕ Adicionar outro item", callback_data="novo_item")],
            [
                InlineKeyboardButton("✏️ Editar Último Item", callback_data="editar_ultimo_item"),
                InlineKeyboardButton("🗑️ Remover Último Item", callback_data="remover_ultimo_item")
            ],
            [InlineKeyboardButton("➡️ Finalizar Pedido", callback_data="finalizar_pedido")],
            [InlineKeyboardButton("💾 Salvar como rascunho", callback_data="salvar_rascunho")]
        ]
        await update.message.reply_text(
            "O que você deseja fazer agora?", 
            reply_markup=InlineKeyboardMarkup(botoes)
        )
        return ITEM_NOVO_OU_FIM

    except ValueError:
        await update.message.reply_text("❌ Preço inválido. Por favor, envie apenas números.")
        return ITEM_PRECO # Permanece no mesmo estado

async def item_remover_ultimo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Remove o último item adicionado à lista e atualiza a mensagem."""
    log_interaction(update)
    query = update.callback_query
    await query.answer()

    itens_do_pedido = context.user_data.get('itens', [])

    if not itens_do_pedido:
        await query.edit_message_text("Não há itens para remover.")
        # Retorna ao mesmo estado para o usuário decidir o que fazer
        return ITEM_NOVO_OU_FIM

    # Remove o último item da lista
    item_removido = itens_do_pedido.pop()
    descricao_removida = item_removido.get('descricao', 'Item desconhecido')
    
    await query.edit_message_text(f"🗑️ O item '{descricao_removida}' foi removido com sucesso!")

    # Aguarda um pouco e mostra as opções novamente
    await asyncio.sleep(2)

    # Reutiliza os botões da função item_preco
    botoes = [
        [InlineKeyboardButton("➕ Adicionar outro item", callback_data="novo_item")],
        [InlineKeyboardButton("✏️ Editar Último Item", callback_data="editar_ultimo_item"),
         InlineKeyboardButton("🗑️ Remover Último Item", callback_data="remover_ultimo_item")],
        [InlineKeyboardButton("➡️ Finalizar Pedido", callback_data="finalizar_pedido")],
        [InlineKeyboardButton("💾 Salvar como rascunho", callback_data="salvar_rascunho")]
    ]
    # Envia uma nova mensagem com as opções
    await query.message.reply_text(
        "O que você deseja fazer agora?",
        reply_markup=InlineKeyboardMarkup(botoes)
    )

    return ITEM_NOVO_OU_FIM


async def item_editar_ultimo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Remove o último item e reinicia o fluxo para adicioná-lo novamente."""
    log_interaction(update)
    query = update.callback_query
    await query.answer()

    itens_do_pedido = context.user_data.get('itens', [])

    if not itens_do_pedido:
        await query.edit_message_text("Não há itens para editar.")
        return ITEM_NOVO_OU_FIM

    # Remove o último item para que ele possa ser inserido novamente
    item_removido = itens_do_pedido.pop()
    descricao_removida = item_removido.get('descricao', 'Item desconhecido')

    context.user_data['item_atual'] = {} # Limpa para o novo item
    
    await query.edit_message_text(
        f"✏️ Ok, vamos corrigir o item '{descricao_removida}'.\n\n"
        "Por favor, informe a <b>nova descrição</b> do item:",
        parse_mode="HTML"
    )
    
    # Retorna a conversa para o estado de pedir a descrição do item
    return ITEM_DESCRICAO

async def item_novo_ou_fim(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    query = update.callback_query
    await query.answer()
    
    if query.data == 'novo_item':
        context.user_data['item_atual'] = {} # Limpa para o próximo item
        await query.edit_message_text("Qual a <b>descrição</b> do próximo item?", parse_mode="HTML")
        return ITEM_DESCRICAO
    
    # --- BLOCO CORRIGIDO ---
    elif query.data == 'finalizar_pedido':
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
        await query.edit_message_text("🔄 Calculando impostos e finalizando o seu pedido...", parse_mode="HTML")
        
        nome_fornecedor = context.user_data.get('fornecedor', '').upper()
        dados_fornecedor = FORNECEDORES.get(nome_fornecedor, {})
        estado_fornecedor = dados_fornecedor.get('estado', '')

        resultado_impostos = call_app_script('_api_calcularImpostosLote', {
            "itens": context.user_data.get('itens', []),
            "estado": estado_fornecedor
        })

        if resultado_impostos.get("status") == "success":
            context.user_data['itens'] = resultado_impostos.get('data', [])
        else:
            erro_backend = resultado_impostos.get('message', 'não especificado')
            await query.message.reply_text(f"⚠️ Não foi possível calcular os impostos (Erro: {erro_backend}). Os valores de imposto serão zerados neste pedido.")
        
        # A lógica do resumo agora está DENTRO do 'elif'
        texto_resumo = "<b>Resumo do Pedido:</b>\n"
        total_geral = 0
        total_icms = 0
        for item in context.user_data.get('itens', []):
            texto_resumo += f"- {item.get('descricao', 'N/A')} ({item.get('quantidade', 0)} {item.get('unidade', 'UN')}) = R$ {item.get('totalItem', 0):.2f}\n"
            total_geral += item.get('totalItem', 0)
            total_icms += item.get('icmsSt', 0)
        
        texto_resumo += f"\n<b>Total ICMS ST: R$ {total_icms:.2f}</b>"
        texto_resumo += f"\n<b>Total Geral: R$ {total_geral:.2f}</b>"
        
        await query.edit_message_text(f"{texto_resumo}\n\nConfirma a criação do pedido? (sim/não)", parse_mode="HTML")
        
        return CRIAR_CONFIRMA

    # --- BLOCO DE RASCUNHO AGORA É ALCANÇÁVEL ---
    elif query.data == 'salvar_rascunho':
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
        await query.edit_message_text("💾 Salvando seu progresso como rascunho...")
        
        resultado = call_app_script("salvar_rascunho", {
            "usuarioId": update.effective_user.id,
            "dadosRascunho": context.user_data
        })
        
        await query.edit_message_text(resultado.get('data', 'Rascunho salvo! Use /rascunhos para continuar mais tarde.'))
        context.user_data.clear()
        
        return ConversationHandler.END
    
async def rascunhos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    """Verifica se o usuário tem um rascunho salvo."""
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
    await update.message.reply_text("🔎 Verificando se há um rascunho salvo para você...")
    
    resultado = call_app_script("carregar_rascunho", {"usuarioId": update.effective_user.id})
    
    if resultado.get("status") == "success":
        # Guarda o rascunho temporariamente para o próximo passo
        context.user_data['rascunho_carregado'] = resultado.get('data')
        
        botoes = [[InlineKeyboardButton("➡️ Continuar Pedido", callback_data="continuar_rascunho")]]
        await update.message.reply_text(
            "✅ Um pedido em andamento foi encontrado! Deseja continuar de onde parou?",
            reply_markup=InlineKeyboardMarkup(botoes)
        )
    else:
        await update.message.reply_text("Nenhum rascunho encontrado. Inicie um novo pedido com /novo_pedido ou novo pedido.")

# Adicione esta nova função de callback
async def continuar_rascunho(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    """Carrega os dados do rascunho e re-entra na conversa."""
    query = update.callback_query
    await query.answer()

    # Pega os dados que o comando /rascunhos guardou
    rascunho = context.user_data.pop('rascunho_carregado', None)
    if not rascunho:
        await query.edit_message_text("Ocorreu um erro ao carregar o rascunho. Tente usar /rascunhos novamente.")
        return ConversationHandler.END

    # Restaura o estado da conversa!
    context.user_data.update(rascunho)

    # --- LÓGICA MELHORADA PARA CRIAR O RESUMO DO RASCUNHO ---
    try:
        texto_resumo = "📝 <b>Pedido restaurado!</b>\n\n"
        
        # Adiciona informações da empresa e fornecedor, se já existirem no rascunho
        if 'empresaInfo' in context.user_data:
            nome_empresa = context.user_data['empresaInfo'].get('empresa', 'N/A')
            texto_resumo += f"<b>Empresa:</b> {nome_empresa}\n"
        if 'fornecedor' in context.user_data:
            texto_resumo += f"<b>Fornecedor:</b> {context.user_data['fornecedor']}\n"
        
        # Adiciona a lista de itens, se houver
        itens = context.user_data.get('itens', [])
        if itens:
            texto_resumo += "\n<b>Itens já adicionados:</b>\n"
            total_parcial = 0
            for item in itens:
                texto_resumo += f"- {item.get('descricao', 'N/A')} ({item.get('quantidade', 0)} {item.get('unidade', 'UN')}) = R$ {item.get('totalItem', 0):.2f}\n"
                total_parcial += item.get('totalItem', 0)
            texto_resumo += f"\n<b>Total Parcial: R$ {total_parcial:.2f}</b>"
        
        texto_resumo += "\n\nVocê pode adicionar mais itens, salvar um novo rascunho ou finalizar o pedido."
        
    except Exception as e:
        # Fallback caso ocorra um erro ao montar o resumo
        logging.error(f"Erro ao montar resumo do rascunho: {e}")
        texto_resumo = "📝 Pedido restaurado! Você pode adicionar mais itens ou finalizar o pedido."
    # --- FIM DA LÓGICA MELHORADA ---
    
    # Manda o usuário de volta para o passo onde ele parou (adicionando itens)
    botoes = [
        [InlineKeyboardButton("➕ Adicionar outro item", callback_data="novo_item")],
        [InlineKeyboardButton("➡️ Finalizar Pedido", callback_data="finalizar_pedido")],
        [InlineKeyboardButton("💾 Salvar Rascunho", callback_data="salvar_rascunho")]
    ]
    await query.edit_message_text(
        texto_resumo,
        reply_markup=InlineKeyboardMarkup(botoes),
        parse_mode="HTML"
    )
    
    # Re-entra na conversa no estado correto para o usuário tomar a próxima ação
    return ITEM_NOVO_OU_FIM # Re-entra na conversa no estado correto
    
async def criar_placa(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Recebe a placa, a valida e, se for válida, pede o nome do veículo."""
    log_interaction(update)
    placa_digitada = update.message.text

    # Chama a função de validação que criamos
    placa_valida_e_normalizada = validar_e_normalizar_placa(placa_digitada)

    if placa_valida_e_normalizada:
        # Se a placa for VÁLIDA:
        context.user_data['placaVeiculo'] = placa_valida_e_normalizada
        await update.message.reply_text("✅ Placa válida! Agora, qual o <b>nome/modelo</b> do veículo?", parse_mode="HTML")
        return CRIAR_NOME_VEICULO # Avança para o próximo estado
    else:
        # Se a placa for INVÁLIDA:
        await update.message.reply_text(
            "❌ Placa inválida. Placas com letras ou números repetidos (ex: 'AAA', '0000') não são permitidas."
        )
        return CRIAR_PLACA

async def criar_nome_veiculo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    log_interaction(update)
    nome_digitado = update.message.text.strip()
    lista_veiculos = context.user_data.get('lista_veiculos', [])

    if not lista_veiculos:
        await update.message.reply_text("⚠️ Não consegui validar o nome do veículo. Aceitando o que foi digitado.")
        context.user_data['nomeVeiculo'] = nome_digitado
        await update.message.reply_text("Qual a <b>descrição</b> do primeiro item?", parse_mode="HTML")
        return ITEM_DESCRICAO

    # Procura por correspondências exatas (ignorando maiúsculas/minúsculas)
    match_exato = next((v for v in lista_veiculos if v.lower() == nome_digitado.lower()), None)
    if match_exato:
        context.user_data['nomeVeiculo'] = match_exato
        await update.message.reply_text(f"✅ Veículo selecionado: <b>{match_exato}</b>.", parse_mode="HTML")
        await update.message.reply_text("Agora, qual a <b>descrição</b> do primeiro item?", parse_mode="HTML")
        return ITEM_DESCRICAO

    # Se não achou, procura por nomes parecidos (similaridade > 80%)
    sugestoes = process.extract(nome_digitado, lista_veiculos, limit=3)
    sugestoes_filtradas = [s[0] for s in sugestoes if s[1] > 80]

    if sugestoes_filtradas:
        botoes = [[InlineKeyboardButton(s, callback_data=f"veiculo_sugerido:{s}")] for s in sugestoes_filtradas]
        await update.message.reply_text(
            "Não encontrei um veículo com esse nome exato. Você quis dizer algum destes?",
            reply_markup=InlineKeyboardMarkup(botoes)
        )
        return CRIAR_NOME_VEICULO # Permanece no mesmo estado
    else:
        await update.message.reply_text(
            "❌ Veículo não encontrado na lista. Por favor, digite o nome novamente ou verifique a grafia."
        )
        return CRIAR_NOME_VEICULO

async def observacoes_opcional_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Lida com a resposta do usuário se ele quer ou não adicionar observações."""
    query = update.callback_query
    await query.answer()
    
    escolha = query.data
    
    if escolha == 'obs_opcional_sim':
        # Usuário quer adicionar observação.
        # O próximo passo será o que definimos em 'proximo_passo_sem_obs'.
        context.user_data['proximo_passo_apos_obs'] = context.user_data.get('proximo_passo_sem_obs')
        await query.edit_message_text("Ok, por favor, digite a <b>observação</b>.", parse_mode="HTML")
        return CRIAR_OBSERVACOES
        
    elif escolha == 'obs_opcional_nao':
        # Usuário NÃO quer adicionar observação.
        # Pulamos para o próximo passo que guardamos na memória.
        proximo_estado = context.user_data.get('proximo_passo_sem_obs')
        
        if proximo_estado == CRIAR_PLACA:
            await query.edit_message_text("Certo. Qual a <b>placa</b> do veículo?", parse_mode="HTML")
            return CRIAR_PLACA
        else: # O padrão é ir para os itens
            context.user_data['itens'] = []
            context.user_data['item_atual'] = {}
            await query.edit_message_text("Ok. Vamos adicionar os itens.\n\nQual a <b>descrição</b> do primeiro item?", parse_mode="HTML")
            return ITEM_DESCRICAO
    
async def criar_observacoes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    """Recebe as observações e avança para a criação de itens."""
    context.user_data['observacoes'] = update.message.text
    
    # Prepara para adicionar o primeiro item
    proximo_estado = context.user_data.get('proximo_passo_apos_obs')

    await update.message.reply_text("✅ Observação salva.")
    
    # Decide para onde ir a seguir com base no que foi salvo >>>
    if proximo_estado == CRIAR_PLACA:
        # Se o próximo passo era pedir a placa, fazemos isso agora.
        await update.message.reply_text("Agora, qual a <b>placa</b> do veículo?", parse_mode="HTML")
        return CRIAR_PLACA
    else:
        # Para todos os outros casos, o padrão é ir para os itens.
        # Prepara para adicionar o primeiro item
        context.user_data['itens'] = []
        context.user_data['item_atual'] = {} 
        
        await update.message.reply_text("Vamos adicionar os itens.\n\n"
                                      "Qual a <b>descrição</b> do primeiro item?", parse_mode="HTML")
        return ITEM_DESCRICAO
    
async def novo_pedido_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    user_name = update.effective_user.first_name
    
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
    # É uma boa prática editar a mensagem de "carregando" em vez de enviar várias
    mensagem_status = await update.message.reply_text("🔄 Sincronizando dados com a planilha, aguarde um instante...")

    try:
        # --- MELHORIA 1: Execução em Paralelo ---
        # Roda as duas tarefas de atualização ao mesmo tempo e espera que ambas terminem.
        # Para rodar uma função síncrona (def) em um ambiente assíncrono (async),
        # usamos asyncio.to_thread.
        resultados_paralelos = await asyncio.gather(
            asyncio.to_thread(atualizar_fornecedores),
            asyncio.to_thread(atualizar_empresas),
            asyncio.to_thread(lambda: call_app_script("listar_veiculos")) # Busca a lista de veículos
        )

        # Processa o resultado da busca de veículos (que é o terceiro item da lista de resultados)
        resultado_veiculos = resultados_paralelos[2]
        logging.info(f"[DEBUG - VEÍCULOS] Resposta da API para 'listar_veiculos': {resultado_veiculos}")

        if resultado_veiculos.get("status") == "success":
            context.user_data['lista_veiculos'] = resultado_veiculos.get('data', [])
            logging.info(f"Lista de veículos carregada com {len(context.user_data['lista_veiculos'])} registros.")
        else:
            # Se a busca falhar, o bot continua funcionando, mas a validação de veículos será pulada.
            context.user_data['lista_veiculos'] = []
            logging.warning("Não foi possível carregar a lista de veículos do backend.")

        # --- MELHORIA 2: Verificação de Erro ---
        # Após a sincronização, verificamos se os dados realmente foram carregados.
        if not EMPRESAS or not FORNECEDORES:
            # Se um deles estiver vazio, algo deu errado na comunicação com a API.
            await mensagem_status.edit_text("❌ Desculpe, não consegui me conectar à planilha de dados agora. Por favor, tente novamente mais tarde.")
            return ConversationHandler.END # Encerra a conversa de forma segura

        # Se tudo deu certo, edita a mensagem de status e continua
        await mensagem_status.edit_text(
            f"Vamos lá, {user_name}! Para começar, me diga para qual <b>empresa</b> é este novo pedido.",
            parse_mode="HTML"
        )
        return CRIAR_EMPRESA

    except Exception as e:
        # Chama o handler de erro principal para registrar o traceback completo do erro
        await erro_handler(update, context)
        
        # O resto do seu código de tratamento de erro continua igual
        logging.error(f"Falha crítica ao sincronizar dados: {e}")
        await mensagem_status.edit_text("❌ Ocorreu um erro crítico ao buscar os dados iniciais. Por favor, avise um administrador.")
        return ConversationHandler.END

async def criar_empresa(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    termo = update.message.text.strip().lower()
    # Busca por parte do nome
    empresas_encontradas = [
        (eid, info) for eid, info in EMPRESAS.items()
        if termo in info["empresa"].lower()
    ]
    if not empresas_encontradas:
        await update.message.reply_text("Não consegui localizar a empresa. Tente novamente.")
        return CRIAR_EMPRESA
    # Se mais de uma, mostra opções
    if len(empresas_encontradas) > 1:
        botoes = [
            [InlineKeyboardButton(info["empresa"], callback_data=f"empresa_escolhida:{eid}")]
            for eid, info in empresas_encontradas
        ]
        await update.message.reply_text(
            "Encontrei mais de uma. Por favor, escolha a empresa correta:",
            reply_markup=InlineKeyboardMarkup(botoes)
        )
        return CRIAR_EMPRESA
    # Se só uma, segue direto
    eid, info = empresas_encontradas[0]
    context.user_data['empresaId'] = eid
    context.user_data['empresaInfo'] = info
    await update.message.reply_text(f"✅ Empresa selecionada: <b>{info['empresa']}</b>.\n\nAgora, informe nome completo ou parte do nome do fornecedor:", parse_mode="HTML")
    return CRIAR_FORNECEDOR

async def empresa_escolhida_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    query = update.callback_query
    await query.answer()
    _, eid = query.data.split(":")
    info = EMPRESAS.get(eid)
    context.user_data['empresaId'] = eid
    context.user_data['empresaInfo'] = info
    await query.edit_message_text("informe nome completo ou parte do nome do fornecedor:")
    return CRIAR_FORNECEDOR

async def _processar_fornecedor_escolhido(update_or_message, context: ContextTypes.DEFAULT_TYPE, nome_fornecedor: str):
    """
    Função central que verifica o grupo do fornecedor e decide o próximo passo da conversa.
    """
    context.user_data['fornecedor'] = nome_fornecedor
    
    #print("\n--- [DEPURAÇÃO] PROCESSANDO ESCOLHA DO FORNECEDOR ---")
    chave_busca = nome_fornecedor.upper()
    #print(f"Buscando pela chave: '{chave_busca}'")

    # Pega os dados completos do fornecedor do nosso mapa global
    dados_fornecedor = FORNECEDORES.get(chave_busca.upper(), {})
    grupo = dados_fornecedor.get('grupo', '').upper()
    # Armazena o grupo para uso futuro (validação do valor do item)
    context.user_data['grupo_fornecedor'] = grupo
    
    nome_para_exibir = dados_fornecedor.get('nome', nome_fornecedor)
    
    texto_resposta = ""
    proximo_estado = None
    texto_confirmacao = f"✅ Fornecedor encontrado: <b>{nome_para_exibir}</b>\n\n"
    # Lógica de decisão
    if grupo == 'LIVRE':
        # Para o grupo LIVRE, as observações são obrigatórias.
        context.user_data['proximo_passo_apos_obs'] = ITEM_DESCRICAO # Depois da obs, vai para os itens
        await update_or_message.reply_text(
            texto_confirmacao + "Para este fornecedor, por favor, adicione uma <b>observação</b> para o pedido.",
            parse_mode="HTML"
        )
        return CRIAR_OBSERVACOES
    else:
        # Para TODOS os outros grupos, a observação é opcional.
        if grupo == 'OBRIGATORIO':
            # Se for obrigatório, o próximo passo sem obs é pedir a PLACA
            context.user_data['proximo_passo_sem_obs'] = CRIAR_PLACA
        else:
            # Para qualquer outro, o próximo passo sem obs é ir para os ITENS
            context.user_data['proximo_passo_sem_obs'] = ITEM_DESCRICAO

        botoes = [[
            InlineKeyboardButton("Sim", callback_data="obs_opcional_sim"),
            InlineKeyboardButton("Não", callback_data="obs_opcional_nao")
        ]]
        await update_or_message.reply_text(
            texto_confirmacao + "Deseja adicionar alguma observação a este pedido?",
            reply_markup=InlineKeyboardMarkup(botoes)
        )
        return CRIAR_OBSERVACOES_OPCIONAL

async def criar_fornecedor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    termo = update.message.text.strip().lower()
    # Lembre-se que FORNECEDORES agora é um dicionário. Buscamos nas chaves.
    fornecedores_encontrados = [
        nome for nome in FORNECEDORES.keys() if termo in nome.lower()
    ]
    
    # ... (a lógica para 0 ou >1 resultados permanece a mesma) ...
    if not fornecedores_encontrados:
        await update.message.reply_text(
            f"🤔 Hmm, não encontrei nenhum fornecedor que contenha '{termo}'.\n\n"
            f"Tente um nome diferente ou use o comando /fornecedores para ver a lista completa."
        )
        return CRIAR_FORNECEDOR
    if len(fornecedores_encontrados) > 1:
        botoes = [[InlineKeyboardButton(nome, callback_data=f"fornecedor_escolhido:{nome}")] for nome in fornecedores_encontrados]
        await update.message.reply_text("Encontrei essa lista abaixo, é algum desses?:", reply_markup=InlineKeyboardMarkup(botoes))
        return CRIAR_FORNECEDOR
    
    
    # Se encontrou apenas 1, processa a escolha
    nome = fornecedores_encontrados[0]
    return await _processar_fornecedor_escolhido(update.message, context, nome)

async def fornecedor_escolhido_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    query = update.callback_query
    await query.answer()
    _, nome = query.data.split(":", 1)
    return await _processar_fornecedor_escolhido(query.message, context, nome)

async def criar_confirma(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    if update.message.text.strip().lower() == 'sim':
        # Envia para o App Script
        params_do_pedido = {
            "empresaId": context.user_data['empresaId'],
            "fornecedor": context.user_data['fornecedor'],
            "itens": context.user_data['itens'],
            "placaVeiculo": context.user_data.get('placaVeiculo', ''),
            "nomeVeiculo": context.user_data.get('nomeVeiculo', ''),
            "observacoes": context.user_data.get('observacoes', '')
            #"status": "rascunho"
        }

        info_do_usuario = {
            "id": update.effective_user.id,
            "first_name": update.effective_user.first_name,
            "username": update.effective_user.username
        }
        resultado = call_app_script(
            "criar_pedido", 
            {"pedido": params_do_pedido, "userInfo": info_do_usuario}
        )
        if resultado.get("status") == "success":
            await update.message.reply_text(resultado['data'], parse_mode="HTML")
            #context.user_data.clear()
            #return ConversationHandler.END
        else:
            erro = resultado.get('message', 'Erro desconhecido')
        await update.message.reply_text(
            f"❌ Ocorreu um erro ao tentar criar o pedido: <i>{erro}</i>\n\n"
            "Por favor, tente novamente. Se o problema persistir, contate o administrador.",
            parse_mode="HTML"
        )
    else:
        await update.message.reply_text("Criação de pedido cancelada.")
        
    context.user_data.clear()
    return ConversationHandler.END

async def calcular_imposto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    """
    Calcula o imposto para um valor e estado fornecidos.
    Uso: /calculoimposto VALOR_ESTADO (Ex: /calculoimposto 1500_SP)
    """
    try:
        # 1. Validação da entrada do usuário
        if not context.args:
            raise ValueError("Argumentos não fornecidos.")
            
        # Pega o primeiro argumento (ex: "1500_SP")
        argumento = context.args[0]
        if '_' not in argumento:
            raise ValueError("Formato inválido. Falta o '_' separador.")

        valor_str, estado = argumento.split('_', 1)
        
        valor = float(valor_str)
        estado = estado.upper() # Garante que o estado esteja em maiúsculas

        if not estado or len(estado) != 2:
            raise ValueError("Estado inválido. Use a sigla de 2 letras.")

    except (ValueError, IndexError) as e:
        logging.warning(f"Erro de validação no /calculoimposto: {e}")
        await update.message.reply_text(
            "😕 Formato inválido. Use:\n"
            "<code>/calculoimposto VALOR_ESTADO</code>\n\n"
            "<b>Exemplo:</b> <code>/calculoimposto 1500_SP</code>",
            parse_mode="HTML"
        )
        return

    # 2. Chamada à API
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
    await update.message.reply_text(f"⏳ Calculando imposto para <b>R$ {valor:.2f}</b> vindo do estado <b>{estado}</b>...", parse_mode="HTML")
    
    resultado = call_app_script(
        'calcular_imposto_simples', 
        {'valor': valor, 'estado': estado}
    )

    # 3. Exibição do resultado
    if resultado.get("status") == "success":
        valor_calculado = resultado.get("valorCalculado", 0)
        aliquota_usada = resultado.get("aliquotaUsada", 0) * 100 # Converte para porcentagem
             
        mensagem = (
            f"📊 <b>Resultado do Cálculo de Imposto</b>\n\n"
            f"<b>Valor Base:</b> R$ {valor:,.2f}\n"
            f"<b>Estado:</b> {estado}\n"
            f"<b>Alíquota Aplicada:</b> {aliquota_usada:.2f}%\n\n"
            f"<b>Valor do Imposto (ICMS ST):</b> <code>R$ {valor_calculado:,.2f}</code>"
        )
        await update.message.reply_text(mensagem, parse_mode="HTML")
    else:
        erro = resultado.get('message', 'Erro desconhecido no servidor.')
        await update.message.reply_text(f"❌ Ocorreu um erro ao calcular: {erro}")

async def dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Envia um dashboard com os dados principais (apenas para admins)."""
    log_interaction(update)
    user_id = str(update.effective_user.id)

    # 1. Verifica se o usuário é um administrador
    if user_id not in ADMIN_CHAT_IDS:
        resposta = "Desculpe, este é um comando exclusivo para administradores."
        await update.message.reply_text(resposta)
        log_interaction(update, bot_response=resposta)
        return

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
    
    # 2. Prepara os filtros com base nos argumentos enviados
    filters = {}
    titulo_partes = [] # Para montar um título dinâmico

    for arg in context.args:
        if "=" not in arg:
            await update.message.reply_text(f"😕 Filtro inválido: '{arg}'. Use o formato `chave=valor`.")
            return
        
        chave, valor = arg.split("=", 1)
        chave = chave.lower().strip()
        valor = valor.strip().replace("_", " ") # Substitui underlines por espaços

        # Mapeia os nomes amigáveis do bot para os nomes esperados pelo backend
        if chave == "periodo":
            titulo_partes.append(f"Período: {valor}")
            if valor == "hoje":
                hoje = datetime.now().strftime('%Y-%m-%d')
                filters['startDate'] = hoje
                filters['endDate'] = hoje
            elif valor == "semana":
                hoje = datetime.now()
                inicio_semana = hoje - timedelta(days=hoje.weekday())
                filters['startDate'] = inicio_semana.strftime('%Y-%m-%d')
                filters['endDate'] = hoje.strftime('%Y-%m-%d')
            elif '-' in valor:
                try:
                    start_str, end_str = valor.split('-')
                    filters['startDate'] = datetime.strptime(start_str, '%d/%m/%Y').strftime('%Y-%m-%d')
                    filters['endDate'] = datetime.strptime(end_str, '%d/%m/%Y').strftime('%Y-%m-%d')
                except ValueError:
                    await update.message.reply_text("😕 Formato de data inválido. Use `DD/MM/AAAA-DD/MM/AAAA`.")
                    return
                
        elif chave == "empresa":
            filters['empresaId'] = valor
            titulo_partes.append(f"Empresa: {valor}")
        elif chave == "fornecedor":
            filters['supplier'] = valor
            titulo_partes.append(f"Fornecedor: {valor}")
        elif chave == "estado":
            filters['state'] = valor
            titulo_partes.append(f"Estado: {valor}")

    # Monta o título final
    titulo = "Dashboard - " + " | ".join(titulo_partes) if titulo_partes else "Dashboard Geral"
    
    # 3. Chama a API do backend com os filtros
    resultado = call_app_script("getDashboardData", filters) 

    # 4. Formata e envia a resposta
    if resultado and resultado.get("status") != 'error':
        total_pedidos = resultado.get("totalPedidos", 0)
        total_gasto = resultado.get("totalGasto", "R$ 0,00")
        ticket_medio = resultado.get("ticketMedio", "R$ 0,00")
        total_icms = resultado.get("totalIcmsSt", "R$ 0,00")

        mensagem = (
            f"📊 <b>{titulo}</b>\n\n"
            f"🛒 <b>Total de Pedidos Aprovados:</b> <code>{total_pedidos}</code>\n"
            f"💰 <b>Valor Total Gasto:</b> <code>{total_gasto}</code>\n"
            f"🎫 <b>Ticket Médio por Pedido:</b> <code>{ticket_medio}</code>\n"
            f"💸 <b>Total de Impostos (ICMS ST):</b> <code>{total_icms}</code>"
        )
        
        if not filters:
            mensagem += (
                "\n\n💡 **Dica:** Você pode adicionar filtros. Ex:\n"
                "`/dashboard periodo=hoje empresa=001`"
            )

        await update.message.reply_text(mensagem, parse_mode="HTML")
    else:
        erro = resultado.get('message', 'Não foi possível obter os dados.')
        await update.message.reply_text(f"❌ Ops! Ocorreu um erro ao buscar os dados: {erro}")

async def cancelar_criacao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log_interaction(update)
    user_name = update.effective_user.first_name
    await update.message.reply_text(f"Tudo bem, {user_name}. A criação do pedido foi cancelada. Se precisar de algo mais, é só chamar!")
    context.user_data.clear()
    return ConversationHandler.END


from flask import Flask
import threading

# Cria o servidor web
app = Flask(__name__)

@app.route('/')
def health_check():
    """Página simples para o UptimeRobot 'visitar'."""
    return "Bot está funcionando perfeitamente!", 200

def run_flask():
    """Roda o servidor Flask em uma porta definida pelo Render."""
    # O Render define a porta através da variável de ambiente PORT
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

async def post_init(application: Application):
    """
    Executa tarefas de inicialização logo após o bot se conectar.
    """
    logging.info("Bot conectado. Sincronizando dados iniciais...")
    # Roda as tarefas de atualização de forma segura
    await asyncio.gather(
        asyncio.to_thread(atualizar_fornecedores), 
        asyncio.to_thread(atualizar_empresas)
    )
    logging.info("Dados iniciais sincronizados com sucesso.")

async def atualizar_admins_job(context: ContextTypes.DEFAULT_TYPE):
    """
    Tarefa agendada que chama a função síncrona para atualizar os admins.
    """
    logging.info("[JobQueue] Executando tarefa agendada: atualizar_admins...")
    # Executa a função síncrona em um contexto que não bloqueia o bot
    await asyncio.to_thread(atualizar_admins)


# --- FUNÇÃO PRINCIPAL ---
def main() -> None:
    """Inicia o bot."""

    application = Application.builder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    
    job_queue = application.job_queue
    
    # Agenda a tarefa para rodar a cada 3600 segundos (1 hora).
    # 'first=20' significa que a primeira execução acontecerá 20 segundos após o bot iniciar.
    job_queue.run_repeating(atualizar_admins_job, interval=3600, first=20)

    async def atualizar_dados_principais_job(context: ContextTypes.DEFAULT_TYPE):
        logging.info("[JobQueue] Executando tarefa agendada: atualizar_dados_principais...")
        await asyncio.gather(
            asyncio.to_thread(atualizar_fornecedores),
            asyncio.to_thread(atualizar_empresas)
        )
    
    job_queue.run_repeating(atualizar_dados_principais_job, interval=3600, first=60)

    application.add_handler(CommandHandler("rascunhos", rascunhos))
    conv_handler = ConversationHandler(
        entry_points=[CallbackQueryHandler(rejeitar_entry, pattern="^rejeitar:.*")],
        states={
            MOTIVO_REJEICAO: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receber_motivo_rejeicao)
            ]
        },
        fallbacks=[CommandHandler("cancelar", cancelar_conversa)],
        conversation_timeout=300
    )
    
    conv_criar_pedido = ConversationHandler(
    entry_points=[
        CommandHandler("novo_pedido", novo_pedido_entry),
        MessageHandler(filters.Regex(re.compile(r'^novo pedido$', re.IGNORECASE)), novo_pedido_entry),
        CallbackQueryHandler(continuar_rascunho, pattern="^continuar_rascunho$")
    ],
    states={
        CRIAR_EMPRESA: [CallbackQueryHandler(empresa_escolhida_callback, pattern="^empresa_escolhida:"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, criar_empresa)],
        CRIAR_FORNECEDOR: [CallbackQueryHandler(fornecedor_escolhido_callback, pattern="^fornecedor_escolhido:"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, criar_fornecedor)],
        
        CRIAR_PLACA: [MessageHandler(filters.TEXT & ~filters.COMMAND, criar_placa)],
        CRIAR_NOME_VEICULO: [MessageHandler(filters.TEXT & ~filters.COMMAND, criar_nome_veiculo)],
        CRIAR_OBSERVACOES_OPCIONAL: [
        CallbackQueryHandler(observacoes_opcional_callback, pattern="^obs_opcional_(sim|nao)$")
        ],
        CRIAR_OBSERVACOES: [MessageHandler(filters.TEXT & ~filters.COMMAND, criar_observacoes)],

        ITEM_DESCRICAO: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_descricao)],
        ITEM_PROD_FORNECEDOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_prod_fornecedor)],
        ITEM_QTD: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_qtd)],
        ITEM_UNIDADE: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_unidade)],
        ITEM_PRECO: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_preco)],
        ITEM_NOVO_OU_FIM: [
            CallbackQueryHandler(item_novo_ou_fim, pattern="^(finalizar_pedido|salvar_rascunho)$"),
            CallbackQueryHandler(item_editar_ultimo, pattern="^editar_ultimo_item$"),
            CallbackQueryHandler(item_remover_ultimo, pattern="^remover_ultimo_item$"),
            CallbackQueryHandler(item_novo_ou_fim, pattern="^novo_item$")
        ],
        CRIAR_CONFIRMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, criar_confirma)],
    },
    fallbacks=[CommandHandler("cancelar", cancelar_criacao)],
    conversation_timeout=600
    )
    
    conv_novo_fornecedor = ConversationHandler(
        entry_points=[CommandHandler("novofornecedor", novo_fornecedor_entry)],
        states={
            CAD_CNPJ: [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_cnpj)],
            CONFIRMAR_CADASTRO: [CallbackQueryHandler(confirmar_cadastro_callback, pattern="^confirmar_cadastro_(sim|nao)$")],
            SELECIONAR_CONDICAO: [CallbackQueryHandler(selecionar_condicao_callback, pattern="^cad_condicao:")],
            SELECIONAR_FORMA: [CallbackQueryHandler(selecionar_forma_callback, pattern="^cad_forma:")],
        },
        fallbacks=[CommandHandler("cancelar", cancelar_conversa)],
    )

    # Este Regex procura por várias saudações no início da mensagem.
    padrao_saudacao = re.compile(r'^\b(oi|ol[aá]|bom dia|boa tarde|boa noite|e a[ií]|opa)\b.*', re.IGNORECASE)
    application.add_handler(MessageHandler(filters.Regex(padrao_saudacao), saudacao))
    # Adicionamos o ConversationHandler para a rejeição de pedidos
    application.add_handler(conv_criar_pedido)
    # Adicionamos o ConversationHandler primeiro, dando a ele prioridade.
    application.add_handler(conv_handler)

    application.add_handler(conv_novo_fornecedor)
    # Handlers para comandos e mensagens de texto gerais
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("buscar", buscar))
    application.add_handler(CommandHandler("dashboard", dashboard))

    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, processar_mensagem_geral))
    application.add_handler(CommandHandler("fornecedores", fornecedores))
    application.add_handler(CallbackQueryHandler(fornecedores_page_callback, pattern="^page_fornecedores:"))
    application.add_handler(CallbackQueryHandler(buscar_fornecedor_callback, pattern="^buscar_fornecedor"))

    application.add_handler(CommandHandler("relatorio_pdf", relatorio_pdf))
    
    application.add_handler(CommandHandler("relatorio_xls", relatorio_xls))
    application.add_handler(CommandHandler("calculoimposto", calcular_imposto))
    # Handler para os outros botões que não fazem parte de uma conversa
    application.add_handler(CallbackQueryHandler(callback_handler, pattern="^(aprovar|detalhes|voltar_para_lista|pdf):.*"))
    
    application.add_error_handler(erro_handler)
    
    #application.add_handler(CallbackQueryHandler(empresa_escolhida_callback, pattern="^empresa_escolhida:"))
    #application.add_handler(CallbackQueryHandler(fornecedor_escolhido_callback, pattern="^fornecedor_escolhido:"))

    application.add_handler(CommandHandler("ajuda", ajuda))
    logging.info("Bot em Python (versão com lista de botões) iniciado com sucesso!")
    application.run_polling()


if __name__ == "__main__":

    # Inicia o servidor web em uma thread separada
    flask_thread = threading.Thread(target=run_flask)
    flask_thread.daemon = True
    flask_thread.start()
    ensure_log_file()
    #atualizar_admins()
    #atualizar_fornecedores(None)
    #atualizar_empresas()
    main()