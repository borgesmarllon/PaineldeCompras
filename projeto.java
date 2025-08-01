<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <base target="_top">
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Portal de Compras</title>
    <!-- Tailwind CSS para um design moderno e responsivo -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Font Awesome para ícones -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.2.0/css/all.min.css">
    <!-- Chart.js para gráficos -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        /* Estilos globais para a aplicação */
        body {
            background-color: #f1f5f9; /* Um cinza claro e suave */
            font-family: 'Inter', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        /* Esconde todas as "telas" por defeito */
        .screen {
            display: none;
        }

        /* Mostra apenas a tela ativa */
        .screen.active {
            display: block;
        }
        
        /* Animação de carregamento (spinner) */
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        .spinner {
            border: 4px solid rgba(0, 0, 0, 0.1);
            border-left-color: #3b82f6; /* Azul */
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
        }
        
        .modal-backdrop {
            display: flex;
        }
        .modal-container {
            transition: opacity 0.3s ease, visibility 0.3s ease;
        }
        
        /* Sistema de Toasts - Notificações modernas */
        .toast-container {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
            pointer-events: none;
            display: flex;
            flex-direction: column;
            gap: 12px;
            max-width: 400px;
        }
        
        .toast {
            pointer-events: auto;
            padding: 16px 20px;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15), 0 4px 6px rgba(0, 0, 0, 0.05);
            backdrop-filter: blur(8px);
            transform: translateX(100%);
            opacity: 0;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
            border-left: 4px solid;
            min-width: 320px;
            max-width: 400px;
        }
        
        .toast.show {
            transform: translateX(0);
            opacity: 1;
        }
        
        .toast.hide {
            transform: translateX(100%);
            opacity: 0;
        }
        
        .toast-success {
            background: linear-gradient(135deg, rgba(34, 197, 94, 0.95) 0%, rgba(21, 128, 61, 0.95) 100%);
            border-left-color: #10b981;
            color: white;
        }
        
        .toast-error {
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.95) 0%, rgba(185, 28, 28, 0.95) 100%);
            border-left-color: #ef4444;
            color: white;
        }
        
        .toast-warning {
            background: linear-gradient(135deg, rgba(245, 158, 11, 0.95) 0%, rgba(180, 83, 9, 0.95) 100%);
            border-left-color: #f59e0b;
            color: white;
        }
        
        .toast-info {
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.95) 0%, rgba(29, 78, 216, 0.95) 100%);
            border-left-color: #3b82f6;
            color: white;
        }
        
        .toast-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 8px;
        }
        
        .toast-title {
            font-weight: 600;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .toast-close {
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            border-radius: 6px;
            width: 24px;
            height: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.2s ease;
            font-size: 12px;
        }
        
        .toast-close:hover {
            background: rgba(255, 255, 255, 0.3);
            transform: scale(1.1);
        }
        
        .toast-message {
            font-size: 13px;
            line-height: 1.5;
            opacity: 0.95;
        }
        
        .toast-progress {
            position: absolute;
            bottom: 0;
            left: 0;
            height: 3px;
            background: rgba(255, 255, 255, 0.4);
            border-radius: 0 0 12px 12px;
            transform-origin: left;
            animation: toast-progress 5s linear forwards;
        }
        
        @keyframes toast-progress {
            from { transform: scaleX(1); }
            to { transform: scaleX(0); }
        }
        
        .toast-icon {
            font-size: 16px;
            margin-right: 2px;
        }
        .modal-panel {
            transition: transform 0.3s ease, opacity 0.3s ease;
        }
        .modal-hidden {
            visibility: hidden;
            opacity: 0;
        }
        .modal-hidden .modal-panel {
            transform: scale(0.9);
            opacity: 0;
        }
        .tag-padrao {
            background-color: #e0f2fe;
            color: #0c4a6e;
            font-weight: 600;
            border: 1px solid #bae6fd;
        }
        .status-ativo { background-color: #dcfce7; color: #166534; }
        .status-inativo { background-color: #fee2e2; color: #991b1b; }
        .tag-padrao {
            background-color: #e0f2fe;
            color: #0c4a6e;
            font-weight: 600;
            border: 1px solid #bae6fd;
        }
        .status-ativo { background-color: #dcfce7; color: #166534; }
        .status-inativo { background-color: #fee2e2; color: #991b1b; }
        
        /* Estilos para a nova tela de busca de pedidos */
        .canceled-row {
            background-color: #fee2e2; /* Red-100 */
            color: #b91c1c; /* Red-700 */
        }
        .canceled-row:hover {
            background-color: #fecaca; /* Red-200 */
        }
        
        /* Spinner de loading */
        .spinner {
            width: 24px;
            height: 24px;
            border: 3px solid #f3f4f6;
            border-top: 3px solid #3b82f6;
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Estilo para a marca d'água */
        .print-watermark {
            position: relative;
        }
        .print-watermark::after {
            content: 'CANCELADO';
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%) rotate(-45deg);
            font-size: 8rem;
            color: rgba(255, 0, 0, 0.15);
            font-weight: bold;
            z-index: 10;
            pointer-events: none;
        }
    </style>
</head>
<body>
    <!-- Container para toasts/notificações -->
    <div id="toast-container" class="toast-container"></div>
    
    <!-- O container principal da nossa aplicação de página única -->
    <div id="app">

        <!-- =============================================== -->
        <!-- TELA DE LOGIN                                   -->
        <!-- =============================================== -->
        <div id="screen-login" class="screen active">
            <div class="min-h-screen flex items-center justify-center p-4">
                <div class="w-full max-w-md bg-white rounded-2xl shadow-xl p-8 space-y-6">
                    <div class="text-center">
                        <h2 class="text-3xl font-bold text-gray-800">Portal de Compras</h2>
                        <p class="text-gray-500 mt-2">Faça login para continuar</p>
                    </div>
                    <div class="space-y-4">
                        <input type="text" id="username" placeholder="Usuário" required autofocus class="w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <input type="password" id="password" placeholder="Senha" required class="w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <select id="selectEmpresaLogin" required class="w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                            <option value="">Primeiro, digite o usuário</option>
                        </select>
                    </div>
                    <button id="loginButton" class="w-full py-3 px-4 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-75">Entrar</button>
                    <p id="message-login" class="text-center text-sm font-medium min-h-[20px]"></p>
                    <div class="text-center">
                        <button class="text-sm text-blue-600 hover:underline" id="createAccountButton">Criar Conta</button>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- =============================================== -->
        <!-- TELA DE CADASTRO DE USUÁRIO                     -->
        <!-- =============================================== -->
        <div id="screen-cadastro" class="screen">
             <div class="max-w-7xl mx-auto">
                  <div class="bg-white rounded-2xl shadow-lg p-6 sm:p-8">
                      <h2 class="text-2xl font-bold text-center text-gray-800 mb-2">Gerenciar Usuários</h2>
                      <p class="text-center text-gray-500 mb-6">Controle de acesso, status e auditoria dos usuários.</p>
                      <p id="userManagementMessage" class="text-center text-sm font-medium min-h-[20px] mb-4 transition-opacity duration-300"></p>
                      <div class="overflow-x-auto">
                          <table class="min-w-full divide-y divide-gray-200">
                              <thead class="bg-gray-50">
                                  <tr>
                                      <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Usuário</th>
                                      <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                                      <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Empresas</th>
                                      <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Ações</th>
                                  </tr>
                              </thead>
                              <tbody id="usersTableBody" class="bg-white divide-y divide-gray-200"></tbody>
                          </table>
                      </div>
                  </div>
              </div>
          </div>

          <!-- Modal de Empresas -->
          <div id="modal-empresas-container" class="fixed inset-0 z-50 flex items-center justify-center p-4 modal-hidden modal-container">
              <div class="absolute inset-0 bg-black/60 modal-overlay"></div>
              <div class="relative w-full max-w-lg bg-white rounded-2xl shadow-xl modal-panel">
                  <div class="p-6">
                      <div class="flex items-start justify-between">
                          <h3 class="text-xl font-bold text-gray-800" id="modal-empresas-title"></h3>
                          <button class="modal-close-btn text-gray-400 hover:text-gray-600 text-2xl leading-none">&times;</button>
                      </div>
                      <div class="mt-6 max-h-72 overflow-y-auto pr-2">
                          <table class="min-w-full">
                              <thead class="bg-gray-50 sticky top-0"><tr><th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Acesso</th><th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Empresa</th><th class="px-4 py-2 text-center text-xs font-medium text-gray-500 uppercase">Padrão</th></tr></thead>
                              <tbody id="empresas-list-body"></tbody>
                          </table>
                      </div>
                  </div>
                  <div class="flex justify-end gap-4 p-4 bg-gray-50 rounded-b-2xl"><button class="modal-cancel-btn px-5 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-100">Cancelar</button><button id="modal-empresas-save-btn" class="px-5 py-2 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700">Salvar</button></div>
              </div>
          </div>
          
          <!-- Modal de Auditoria -->
          <div id="modal-audit-container" class="fixed inset-0 z-50 flex items-center justify-center p-4 modal-hidden modal-container">
              <div class="absolute inset-0 bg-black/60 modal-overlay"></div>
              <div class="relative w-full max-w-lg bg-white rounded-2xl shadow-xl modal-panel">
                  <div class="p-6">
                      <div class="flex items-start justify-between">
                          <h3 class="text-xl font-bold text-gray-800" id="modal-audit-title"></h3>
                          <button class="modal-close-btn text-gray-400 hover:text-gray-600 text-2xl leading-none">&times;</button>
                      </div>
                      <div id="modal-audit-body" class="mt-6 space-y-4"></div>
                  </div>
                  <div class="flex justify-end gap-4 p-4 bg-gray-50 rounded-b-2xl"><button class="modal-cancel-btn px-5 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-100">Fechar</button></div>
              </div>
          </div>

        <!-- =============================================== -->
        <!-- TELA PRINCIPAL (após o login)                   -->
        <!-- =============================================== -->
        <div id="screen-main" class="screen">
            <!-- Barra de Navegação Superior -->
            <nav class="bg-white shadow-md sticky top-0 z-40">
                <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
                    <div class="flex items-center justify-between h-16">
                        <div class="flex items-center">
                            <span class="text-2xl font-bold text-blue-600">Portal de Compras</span>
                            <div id="empresaInfo" class="ml-6 hidden md:block text-sm text-gray-600 border-l-2 border-gray-200 pl-4"></div>
                        </div>
                        <div class="flex items-center">
                            <div class="welcome-message text-sm text-gray-700 mr-4 hidden sm:block">
                                Olá, <span id="nome" class="font-semibold"></span> (<span id="perfil"></span>)
                            </div>
                            <button id="logoutButton" class="flex items-center px-3 py-2 bg-red-500 text-white text-sm font-medium rounded-md hover:bg-red-600 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-red-500">
                                <i class="fas fa-sign-out-alt mr-2"></i>Sair
                            </button>
                        </div>
                    </div>
                </div>
            </nav>

            <!-- Área de Conteúdo Principal -->
            <main id="main-content-area" class="p-4 sm:p-6 lg:p-8">
                <!-- O conteúdo das outras "telas" será injetado aqui -->
            </main>
        </div>
        
        <!-- MODAL GLOBAL DE AVISOS (para substituir alerts) -->
        <div id="global-modal" class="fixed inset-0 bg-black bg-opacity-50 hidden justify-center items-center p-4 z-50">
            <div class="bg-white rounded-xl shadow-2xl w-full max-w-md p-6 text-center space-y-4">
                <h3 id="modal-title" class="text-xl font-bold text-gray-800">Aviso</h3>
                <p id="modal-message" class="text-gray-600">Sua mensagem aparecerá aqui.</p>
                <button id="modal-close-btn" class="px-6 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700">OK</button>
            </div>
        </div>
        
        <!-- MODAL DE REDEFINIÇÃO DE SENHA -->
        <div id="modalRedefinirSenha" class="fixed inset-0 bg-black bg-opacity-50 hidden items-center justify-center p-4 z-50">
            <div class="bg-white rounded-xl shadow-2xl w-full max-w-md p-6 space-y-4">
                <h3 class="text-xl font-bold text-gray-800 text-center">Redefinir senha de <span id="modalUsuario" class="text-blue-600"></span></h3>
                <input id="novaSenhaInput" type="password" placeholder="Nova senha" class="w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg">
                <p id="modalSenhaMsg" class="text-center text-sm font-medium min-h-[20px]"></p>
                <div class="flex space-x-4">
                    <button onclick="fecharModalRedefinirSenha()" class="w-full py-2 px-4 bg-gray-500 text-white font-semibold rounded-lg hover:bg-gray-600">Cancelar</button>
                    <button onclick="confirmarRedefinirSenha()" class="w-full py-2 px-4 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700">Confirmar</button>
                </div>
            </div>
        </div>

    </div>

    <!-- TODO O SEU JAVASCRIPT, ADAPTADO PARA O MODELO SPA -->
    <script>
        // ===============================================
        // SETUP INICIAL E CONTROLO DE TELAS (SPA)
        // ===============================================
        const appDiv = document.getElementById('app');
        const mainContentArea = document.getElementById('main-content-area');

        function showScreen(screenId) {
            document.querySelectorAll('.screen').forEach(screen => {
                screen.classList.remove('active');
            });
            const targetScreen = document.getElementById(screenId);
            if (targetScreen) {
                targetScreen.classList.add('active');
                
                // Chamar função de setup específica para cada tela
                if (screenId === 'screen-login') {
                    setTimeout(() => setupLoginScreen(), 100);
                } else if (screenId === 'screen-cadastro') {
                    setTimeout(() => setupCadastroScreen(), 100);
                }
                
                console.log('Tela', screenId, 'ativada com sucesso');
            } else {
                console.error(`Tela com ID "${screenId}" não encontrada.`);
            }
        }

        function loadPageContent(pageName, callback, ...args) {
            mainContentArea.innerHTML = `<div class="flex justify-center items-center p-10"><div class="spinner"></div></div>`;
            const pageHtml = getPageTemplate(pageName, ...args);

            if (pageHtml) {
                mainContentArea.innerHTML = pageHtml;
                if (callback && typeof callback === 'function') {
                    setTimeout(() => callback(...args), 50);
                }
            } else {
                 mainContentArea.innerHTML = `<div class="text-center text-red-500 font-bold p-10">Erro: Conteúdo para a página "${pageName}" não encontrado.</div>`;
            }
        }
        
        function getPageTemplate(pageName, ...args) {
            const templates = {
                Menu: `
                    <div class="max-w-7xl mx-auto grid grid-cols-1 lg:grid-cols-3 gap-8">
                        <div class="lg:col-span-1 space-y-8">
                            <div class="bg-white p-6 rounded-xl shadow-lg">
                                <h3 class="text-xl font-bold text-gray-800 mb-4">Ações da Empresa</h3>
                                <div class="space-y-4">
                                    <div>
                                        <label for="selectEmpresa" class="block text-sm font-medium text-gray-700 mb-1">Trocar de Empresa:</label>
                                        <select id="selectEmpresa" class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-md"></select>
                                    </div>
                                    <button id="btnSelecionarEmpresa" class="w-full text-center px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700">Confirmar Troca</button>
                                </div>
                            </div>
                            <div class="bg-white p-6 rounded-xl shadow-lg">
                                <h3 class="text-xl font-bold text-gray-800 mb-4">Menu Principal</h3>
                                <div id="menuOptionsGroup" class="space-y-3"></div>
                            </div>
                        </div>
                        <div id="company-dashboard-container" class="lg:col-span-2 space-y-6">
                             <div class="bg-white p-6 rounded-xl shadow-lg">
                                <h3 class="text-xl font-bold text-gray-800 mb-4">Dashboard Rápido</h3>
                                <div id="dashboard-content-area">
                                    <div class="text-center py-10 text-gray-500">
                                        <div class="spinner mx-auto"></div>
                                        <p class="mt-4">A carregar...</p>
                                    </div>
                                </div>
                             </div>
                        </div>
                    </div>`,
                GerenciarUsuarios: `
                    <div class="bg-white rounded-2xl shadow-lg p-6 sm:p-8">
                        <h2 class="text-2xl font-bold text-center text-gray-800 mb-2">Gerenciar Usuários</h2>
                        <p class="text-center text-gray-500 mb-6">Controle de acesso, status e auditoria dos usuários.</p>
                        <p id="userManagementMessage" class="text-center text-sm font-medium min-h-[20px] mb-4 transition-opacity duration-300"></p>
                        <div class="overflow-x-auto">
                            <table class="min-w-full divide-y divide-gray-200">
                                <thead class="bg-gray-50">
                                    <tr>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Usuário</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Empresas</th>
                                        <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Ações</th>
                                    </tr>
                                </thead>
                                <tbody id="usersTableBody" class="bg-white divide-y divide-gray-200"></tbody>
                            </table>
                        </div>
                        <div class="mt-6 text-center">
                            <button id="backToMenuFromUsersButton" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">Voltar ao Início</button>
                        </div>
                    </div>
                    
                    <!-- Modal de Empresas -->
                    <div id="modal-empresas-container" class="fixed inset-0 z-50 flex items-center justify-center p-4 modal-hidden modal-container">
                        <div class="absolute inset-0 bg-black/60 modal-overlay"></div>
                        <div class="relative w-full max-w-lg bg-white rounded-2xl shadow-xl modal-panel">
                            <div class="p-6">
                                <div class="flex items-start justify-between">
                                    <h3 class="text-xl font-bold text-gray-800" id="modal-empresas-title"></h3>
                                    <button class="modal-close-btn text-gray-400 hover:text-gray-600 text-2xl leading-none">&times;</button>
                                </div>
                                <div class="mt-6 max-h-72 overflow-y-auto pr-2">
                                    <table class="min-w-full">
                                        <thead class="bg-gray-50 sticky top-0">
                                            <tr>
                                                <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Acesso</th>
                                                <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Empresa</th>
                                                <th class="px-4 py-2 text-center text-xs font-medium text-gray-500 uppercase">Padrão</th>
                                            </tr>
                                        </thead>
                                        <tbody id="empresas-list-body"></tbody>
                                    </table>
                                </div>
                            </div>
                            <div class="flex justify-end gap-4 p-4 bg-gray-50 rounded-b-2xl">
                                <button class="modal-cancel-btn px-5 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-100">Cancelar</button>
                                <button id="modal-empresas-save-btn" class="px-5 py-2 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700">Salvar</button>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Modal de Auditoria -->
                    <div id="modal-audit-container" class="fixed inset-0 z-50 flex items-center justify-center p-4 modal-hidden modal-container">
                        <div class="absolute inset-0 bg-black/60 modal-overlay"></div>
                        <div class="relative w-full max-w-lg bg-white rounded-2xl shadow-xl modal-panel">
                            <div class="p-6">
                                <div class="flex items-start justify-between">
                                    <h3 class="text-xl font-bold text-gray-800" id="modal-audit-title"></h3>
                                    <button class="modal-close-btn text-gray-400 hover:text-gray-600 text-2xl leading-none">&times;</button>
                                </div>
                                <div id="modal-audit-body" class="mt-6 space-y-4"></div>
                            </div>
                            <div class="flex justify-end gap-4 p-4 bg-gray-50 rounded-b-2xl">
                                <button class="modal-cancel-btn px-5 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-100">Fechar</button>
                            </div>
                        </div>
                    </div>`,
                CadastroFornecedor: `
                    <div class="max-w-2xl mx-auto bg-white rounded-2xl shadow-xl p-8 space-y-6">
                        <h2 class="text-2xl font-bold text-center text-gray-800">Cadastro de Fornecedor</h2>
                        <form id="form-fornecedor" autocomplete="off" class="space-y-6">
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                                <div class="md:col-span-1">
                                    <label for="codigo" class="block text-sm font-medium text-gray-700">Código</label>
                                    <input type="text" id="codigo" name="codigo" readonly class="mt-1 block w-full px-3 py-2 bg-gray-200 border border-gray-300 rounded-md shadow-sm cursor-not-allowed">
                                </div>
                                <div class="md:col-span-2">
                                    <label for="cnpj" class="block text-sm font-medium text-gray-700">CNPJ</label>
                                    <div class="mt-1 flex rounded-md shadow-sm">
                                        <input type="text" id="cnpj" name="cnpj" required placeholder="Digite o CNPJ" class="flex-1 block w-full min-w-0 rounded-none rounded-l-md px-3 py-2 border-gray-300">
                                        <button type="button" id="consultarCnpjBtn" class="inline-flex items-center px-4 py-2 border border-l-0 border-gray-300 rounded-r-md bg-gray-50 text-sm font-medium text-gray-700 hover:bg-gray-100">
                                            <i class="fas fa-search"></i>
                                        </button>
                                    </div>
                                </div>
                            </div>
                            <div>
                                <label for="razao" class="block text-sm font-medium text-gray-700">Razão Social</label>
                                <input type="text" id="razao" name="razao" required placeholder="Digite a Razão Social" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm">
                            </div>
                            <div>
                                <label for="fantasia" class="block text-sm font-medium text-gray-700">Nome Fantasia</label>
                                <input type="text" id="fantasia" name="fantasia" required placeholder="Digite o Nome Fantasia" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm">
                            </div>
                            <div>
                                <label for="endereco" class="block text-sm font-medium text-gray-700">Endereço</label>
                                <input type="text" id="endereco" name="endereco" required placeholder="Endereço completo do fornecedor" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm">
                            </div>
                            <div>
                                <label for="estado" class="block text-sm font-medium text-gray-700">Estado (UF)</label>
                                <select id="estado" name="estado" required class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 rounded-md">
                                    <option value="">Carregando estados...</option>
                                </select>
                            </div>
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                                <div>
                                    <label for="condicao" class="block text-sm font-medium text-gray-700">Condição de Pagamento</label>
                                    <select id="condicao" name="condicao" required class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 rounded-md"></select>
                                </div>
                                <div>
                                    <label for="forma" class="block text-sm font-medium text-gray-700">Forma de Pagamento</label>
                                    <select id="forma" name="forma" required class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 rounded-md"></select>
                                </div>
                            </div>
                            <p id="msgFornecedor" class="text-center text-sm font-medium min-h-[20px]"></p>
                            <div class="flex justify-end space-x-4 pt-4 border-t border-gray-200">
                                <button type="button" id="cancelButton" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">Cancelar</button>
                                <button type="button" id="saveButton" class="px-6 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700">Salvar</button>
                            </div>
                        </form>
                    </div>`,
                Pedido: `
                    <style>
                        /* Faz o card de totais "flutuar" na lateral em telas grandes */
                        .sticky-summary {
                            position: sticky;
                            top: 2rem;
                        }
                        
                        /* Estilo para campos obrigatórios */
                        .campo-obrigatorio {
                            border-left: 4px solid #f59e0b;
                        }
                        
                        /* Indicador de rascunho */
                        .draft-indicator {
                            background: linear-gradient(45deg, #fbbf24, #f59e0b);
                        }
                    </style>
                    
                    <div class="max-w-7xl mx-auto">
                        <header class="mb-6">
                            <div class="flex items-center justify-between">
                                <div>
                                    <h1 id="page-title" class="text-3xl font-bold text-gray-800">Novo Pedido de Compra</h1>
                                    <p id="page-subtitle" class="text-gray-500">Preencha os dados e adicione os itens manualmente.</p>
                                </div>
                                <div id="draft-status" class="hidden px-4 py-2 bg-yellow-100 text-yellow-800 rounded-lg border border-yellow-300">
                                    <i class="fas fa-edit mr-2"></i>Rascunho
                                </div>
                            </div>
                            <p id="msgPedido" class="text-center text-sm font-medium min-h-[20px] mt-2"></p>
                        </header>

                        <main class="grid grid-cols-1 lg:grid-cols-3 gap-8">
                            <!-- Coluna Principal (Esquerda) -->
                            <div class="lg:col-span-2 space-y-8">
                                <!-- Dados do Pedido e Fornecedor -->
                                <div class="bg-white p-6 rounded-xl shadow-lg">
                                    <h3 class="font-bold text-lg text-gray-700 mb-4">1. Informações do Pedido</h3>
                                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                                        <div>
                                            <label for="numeroPedido" class="block text-sm font-medium text-gray-700">Número do Pedido</label>
                                            <input type="text" id="numeroPedido" name="numeroPedido" placeholder="Será gerado automaticamente" readonly class="mt-1 block w-full px-3 py-2 bg-gray-100 text-gray-500 rounded-md cursor-not-allowed">
                                        </div>
                                        <div>
                                            <label for="dataPedido" class="block text-sm font-medium text-gray-700">Data do Pedido</label>
                                            <input type="date" id="dataPedido" name="dataPedido" readonly class="mt-1 block w-full px-3 py-2 bg-gray-100 text-gray-500 rounded-md cursor-not-allowed">
                                        </div>
                                        <div>
                                            <label class="block text-sm font-medium text-gray-700">Estado (UF)</label>
                                            <p id="fornecedor-estado" class="mt-1 block w-full px-3 py-2 bg-gray-100 text-gray-500 rounded-md">Selecione</p>
                                        </div>
                                    </div>
                                    <div class="grid grid-cols-1 gap-4">
                                        <div>
                                            <label for="fornecedorPedido" class="block text-sm font-medium text-gray-700">Fornecedor <span class="text-red-500">*</span></label>
                                            <select id="fornecedorPedido" name="fornecedorPedido" required class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-md">
                                                <option value="">-- Selecione um Fornecedor --</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>

                            <!-- Informações Adicionais -->
                                <div class="bg-white p-6 rounded-xl shadow-lg">
                                    <h3 class="font-bold text-lg text-gray-700 mb-4">2. Informações Adicionais</h3>
                                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-4">
                                        <div>
                                            <label for="nomeVeiculo" class="block text-sm font-medium text-gray-700">
                                                Nome do Veículo <span id="veiculo-required" class="text-red-500 hidden">*</span>
                                            </label>
                                            <div class="mt-1 flex rounded-md shadow-sm">
                                                <select id="nomeVeiculo" name="nomeVeiculo" class="flex-1 block w-full min-w-0 rounded-none rounded-l-md px-3 py-2 border-gray-300">
                                                    <option value="">-- Selecione um Veículo --</option>
                                                </select>
                                                <button type="button" id="addVeiculoButton" title="Adicionar Novo Veículo" class="inline-flex items-center px-3 py-2 border border-l-0 border-gray-300 rounded-r-md bg-green-500 text-white hover:bg-green-600">
                                                    <i class="fas fa-plus"></i>
                                                </button>
                                            </div>
                                        </div>
                                        <div>
                                            <label for="placaVeiculo" class="block text-sm font-medium text-gray-700">
                                                Placa <span id="placa-required" class="text-red-500 hidden">*</span>
                                            </label>
                                            <input type="text" id="placaVeiculo" name="placaVeiculo" placeholder="Ex: ABC-1234" class="mt-1 block w-full px-3 py-2 bg-gray-50 border-gray-300 rounded-md">
                                        </div>
                                    </div>
                                    <div>
                                        <label for="observacoesPedido" class="block text-sm font-medium text-gray-700">
                                            Observações do Pedido <span id="obs-required" class="text-red-500 hidden">*</span>
                                        </label>
                                        <textarea id="observacoesPedido" name="observacoesPedido" rows="3" placeholder="Adicione observações importantes para o pedido..." class="mt-1 block w-full px-3 py-2 bg-gray-50 border-gray-300 rounded-md"></textarea>
                                    </div>
                                </div>

                                <!-- Seção dos Itens existentes foi removida para não conflitar com o layout atual -->

                                <!-- Formulário para Adicionar Itens -->
                                <div class="bg-white p-6 rounded-xl shadow-lg">
                                    <h3 class="font-bold text-lg text-gray-700 mb-4">3. Adicionar Itens ao Pedido</h3>
                                    <div class="space-y-4">
                                        <div class="grid grid-cols-1 md:grid-cols-6 gap-4 items-end">
                                            <div class="md:col-span-2">
                                                <label for="itemDescricao" class="block text-sm font-medium text-gray-700">Descrição do Item</label>
                                                <input type="text" id="itemDescricao" name="itemDescricao" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm" placeholder="Ex: Óleo Motor 15W40">
                                            </div>
                                            <div>
                                                <label for="itemQuantidade" class="block text-sm font-medium text-gray-700">Qtd.</label>
                                                <input type="number" id="itemQuantidade" name="itemQuantidade" value="1" min="0.01" step="0.01" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm">
                                            </div>
                                            <div>
                                                <label for="itemUnidade" class="block text-sm font-medium text-gray-700">Unid.</label>
                                                <select id="itemUnidade" name="itemUnidade" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm">
                                                    <option value="UN">UN</option>
                                                    <option value="MT">MT</option>
                                                    <option value="L">L</option>
                                                    <option value="RL">RL</option>
                                                    <option value="CX">CX</option>
                                                </select>
                                            </div>
                                            <div class="md:col-span-2">
                                                <label for="itemPrecoUnitario" class="block text-sm font-medium text-gray-700">Valor Unit. (R$)</label>
                                                <input type="text" id="itemPrecoUnitario" name="itemPrecoUnitario" placeholder="0,00" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm text-right">
                                            </div>
                                        </div>
                                        <div class="flex justify-between items-center pt-4 border-t">
                                            <span class="text-sm text-gray-500">Subtotal do item: <strong id="live-subtotal" class="text-gray-800">R$ 0,00</strong></span>
                                            <button type="button" id="addItemButton" class="bg-blue-600 text-white font-semibold py-2 px-4 rounded-lg shadow-md hover:bg-blue-700 flex items-center justify-center">
                                                <i class="fas fa-plus mr-2"></i> Adicionar ao Pedido
                                            </button>
                                        </div>
                                    </div>
                                </div>
                                
                                <!-- Tabela de Itens Adicionados -->
                                <div class="bg-white rounded-xl shadow-lg">
                                    <h3 class="font-bold text-lg text-gray-700 p-6">4. Itens no Pedido</h3>
                                    <div class="overflow-x-auto">
                                        <table class="min-w-full">
                                            <thead class="bg-gray-50">
                                                <tr>
                                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">Descrição</th>
                                                    <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase">Qtd.</th>
                                                    <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase">Unid.</th>
                                                    <th class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">Vl. Unit.</th>
                                                    <th class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">Subtotal</th>
                                                    <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase">Ações</th>
                                                </tr>
                                            </thead>
                                            <tbody id="itensTableBody" class="bg-white divide-y divide-gray-200">
                                                <tr id="no-items-row"><td colspan="6" class="text-center text-gray-500 py-10">Nenhum item adicionado.</td></tr>
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>    

                            <!-- Coluna de Resumo (Direita) -->
                            <div class="lg:col-span-1">
                                <div class="bg-white p-6 rounded-xl shadow-lg sticky-summary space-y-4">
                                    <h3 class="font-bold text-lg text-gray-700 border-b pb-3">Resumo Financeiro</h3>
                                    
                                    <div class="flex justify-between text-gray-600">
                                        <span>Soma dos Itens</span>
                                        <span id="totalGeral">R$ 0,00</span>
                                    </div>
                                    <div class="flex justify-between text-gray-600">
                                        <span>Impostos (ICMS)</span>
                                        <span id="summary-taxes">R$ 0,00</span>
                                    </div>
                                    
                                    <div class="border-t-2 border-dashed pt-4">
                                        <div class="flex justify-between items-center font-bold text-xl text-gray-800">
                                            <span>Total Geral</span>
                                            <span id="summary-total">R$ 0,00</span>
                                        </div>
                                    </div>

                                    <div class="pt-6 space-y-3">
                                        <button id="salvarRascunhoButton" class="w-full bg-yellow-500 text-white font-bold py-3 px-4 rounded-lg shadow-md hover:bg-yellow-600">
                                            <i class="fas fa-save mr-2"></i>Salvar Rascunho
                                        </button>
                                        <button id="salvarPedidoButton" class="w-full bg-green-600 text-white font-bold py-3 px-4 rounded-lg shadow-md hover:bg-green-700">
                                            <i class="fas fa-check mr-2"></i>Finalizar e Salvar Pedido
                                        </button>
                                        <button id="cancelarPedidoButton" class="w-full bg-gray-200 text-gray-700 font-bold py-3 px-4 rounded-lg hover:bg-gray-300">
                                            <i class="fas fa-times mr-2"></i>Cancelar
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </main>
                    </div>`,
                PedidoSalvo: `
                    <div class="text-center p-8 bg-white rounded-2xl shadow-xl space-y-6 max-w-lg mx-auto">
                        <i class="fas fa-check-circle text-6xl text-green-500"></i>
                        <h2 class="text-2xl font-bold text-gray-800">Pedido Salvo!</h2>
                        <p class="text-gray-600">O pedido <strong>${args[0]}</strong> foi salvo com sucesso.</p>
                        <div class="flex justify-center space-x-4">
                            <button id="voltarMenuPrincipalButton" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg hover:bg-gray-700">Voltar ao Início</button>
                            <button id="printPedidoButton" class="px-6 py-2 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700">
                                <i class="fas fa-print mr-2"></i>Imprimir Pedido
                            </button>
                        </div>
                    </div>`,
                PedidoImpressao: `
                    <div>
                        <style>
                            body { font-family: 'Arial', sans-serif; margin: 0; padding: 0; background-color: #fff; -webkit-print-color-adjust: exact; color-adjust: exact; }
                            .print-container { width: 210mm; min-height: 297mm; margin: 10mm auto; border: 1px solid #ccc; box-shadow: 0 0 5px rgba(0,0,0,0.1); background-color: #fff; padding: 15mm; box-sizing: border-box; }
                            .header-section { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 1px solid #000; padding-bottom: 10px; margin-bottom: 15px; }
                            .company-info { display: flex; align-items: flex-start; flex: 2; text-align: left; }
                            .company-logo { width: 60px; height: 60px; margin-right: 15px; border: 1px solid #eee; padding: 5px; }
                            .company-details { font-size: 0.8em; line-height: 1.2; }
                            .company-details strong { font-size: 1.2em; display: block; margin-bottom: 3px; }
                            .company-details p { margin: 0; }
                            .order-info { flex: 1; display: flex; flex-direction: column; align-items: flex-end; font-size: 0.9em; gap: 5px; }
                            .info-box { border: 1px solid #000; padding: 5px 10px; text-align: center; width: 100%; box-sizing: border-box; background-color: #f0f0f0; }
                            .info-box span:first-child { display: block; font-weight: bold; margin-bottom: 2px; font-size: 0.9em; }
                            .info-box span:last-child { font-size: 1.1em; font-weight: bold; color: #0056b3; }
                            .supplier-section { display: flex; justify-content: space-between; border-bottom: 1px solid #000; padding-bottom: 5px; margin-bottom: 10px; font-size: 0.85em; }
                            .supplier-details, .supplier-meta { flex: 1; text-align: left; }
                            .supplier-section p { margin: 2px 0; }
                            .supplier-details { padding-right: 15px; }
                            .vehicle-section { border-bottom: 1px solid #000; padding-bottom: 5px; margin-bottom: 10px; font-size: 0.85em; text-align: left; }
                            .vehicle-section p { margin: 2px 0; display: inline-block; margin-right: 20px; }
                            .items-table-section { margin-bottom: 15px; }
                            table { width: 100%; border-collapse: collapse; font-size: 0.8em; }
                            th, td { border: 1px solid #000; padding: 6px 8px; text-align: left; vertical-align: top; }
                            th { background-color: #f0f0f0; font-weight: bold; text-align: center; }
                            td { white-space: normal; word-wrap: break-word; }
                            .total-notes-section { border: 1px solid #000; padding: 10px; margin-top: 20px; margin-bottom: 20px; background-color: #f5f5f5; }
                            .total-value { font-size: 1.1em; font-weight: bold; text-align: right; margin-bottom: 10px; }
                            .notes { font-size: 0.8em; text-align: center; line-height: 1.3; }
                            .attention-text { font-weight: bold; color: #dc3545; }
                            .signature-section { text-align: center; margin-top: 30px; font-size: 0.9em; font-weight: bold; }
                            .signature-section p { margin: 2px 0; }
                            @media print { .print-container { border: none; box-shadow: none; margin: 0; width: 100%; min-height: auto; padding: 0; } body { margin: 0; padding: 0; } }
                        </style>
                        <div class="print-container">
                            <div class="header-section">
                                <div class="company-info">
                                    <img src="https://placehold.co/100x100/eeeeee/black?text=LOGO" alt="Logo da Empresa" class="company-logo">
                                    <div class="company-details">
                                        <strong id="empresaNomePrint"></strong><p id="empresaEnderecoPrint"></p><p id="empresaCidadeUfPrint"></p>
                                        <p>CNPJ: <span id="empresaCnpjPrint"></span></p><p id="empresaEmailPrint"></p><p id="empresaTelefonePrint"></p>
                                    </div>
                                </div>
                                <div class="order-info">
                                    <div class="info-box date-box"><span>DATA DA EMISSÃO</span><span id="dataEmissaoPrint"></span></div>
                                    <div class="info-box order-number-box"><span>NÚMERO DO PEDIDO</span><span id="numeroPedidoPrint"></span></div>
                                </div>
                            </div>
                            <div class="supplier-section">
                                <div class="supplier-details">
                                    <p><strong>FORNECEDOR:</strong> <span id="fornecedorPrint"></span></p><p><strong>ENDEREÇO:</strong> <span id="enderecoFornecedorPrint"></span></p><p><strong>FORMA PGTO:</strong> <span id="formaPagamentoPrint"></span></p>
                                </div>
                                <div class="supplier-meta">
                                    <p><strong>CNPJ:</strong> <span id="cnpjFornecedorPrint"></span></p><p><strong>CONDIÇÃO PGTO:</strong> <span id="condicaoPagamentoPrint"></span></p>
                                </div>
                            </div>
                            <div class="vehicle-section">
                                <p><strong>PLACA:</strong> <span id="placaVeiculoPrint"></span></p><p><strong>VEÍCULO:</strong> <span id="nomeVeiculoPrint"></span></p>
                            </div>
                            <div class="items-table-section">
                                <table>
                                    <thead><tr><th style="width: 5%;">CÓD.</th><th style="width: 40%;">DESCRIÇÃO</th><th style="width: 15%;">UNIDADE</th><th style="width: 10%;">QTD.</th><th style="width: 15%;">VALOR UNIT.</th><th style="width: 15%;">SUBTOTAL</th></tr></thead>
                                    <tbody id="printItemsTableBody"></tbody>
                                </table>
                            </div>
                            <div class="total-notes-section">
                                <div class="total-value"><strong>VALOR TOTAL DA NOTA:</strong> <span id="totalGeralPrint"></span></div>
                                <div class="notes"><p id="observacoesPedidoPrint" class="note-text"></p><p class="attention-text">Atenção: Qualquer alteração só pode ser realizada mediante autorização prévia, sob risco de não pagamento.</p></div>
                            </div>
                            <div class="signature-section">
                                <p><strong id="usuarioLogadoPrint"></strong></p><p id="funcaoUsuarioLogadoPrint"></p>
                            </div>
                        </div>
                    </div>`,
                BuscarPedido: `
                    <div class="max-w-7xl mx-auto bg-white p-6 sm:p-8 rounded-2xl shadow-lg">
                        <h2 class="text-2xl font-bold text-gray-800 mb-6">Buscar Ordem de Compra</h2>

                        <!-- Filtros de Busca -->
                        <div class="space-y-4">
                            <!-- Filtro Principal -->
                            <div class="flex flex-col sm:flex-row gap-4">
                                <input type="text" id="mainSearch" placeholder="Buscar por Nº do Pedido ou Fornecedor..." class="flex-grow w-full px-4 py-2 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                                <button id="searchButton" class="w-full sm:w-auto px-6 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700">
                                    <i class="fas fa-search mr-2"></i>Buscar
                                </button>
                            </div>
                            <!-- Botão para Filtros Avançados -->
                            <div>
                                <button id="toggleAdvancedFilters" class="text-sm text-blue-600 hover:underline">
                                    Filtros Avançados <i class="fas fa-chevron-down ml-1 transition-transform"></i>
                                </button>
                            </div>
                            <!-- Filtros Avançados (escondidos por defeito) -->
                            <div id="advancedFilters" class="hidden pt-4 border-t border-gray-200">
                                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                                    <input type="date" id="dateStart" title="Data Inicial" class="px-4 py-2 bg-gray-50 border border-gray-300 rounded-lg">
                                    <input type="date" id="dateEnd" title="Data Final" class="px-4 py-2 bg-gray-50 border border-gray-300 rounded-lg">
                                    <input type="text" id="plateSearch" placeholder="Placa do Veículo" class="px-4 py-2 bg-gray-50 border border-gray-300 rounded-lg">
                                    <input type="text" id="userSearch" placeholder="Utilizador que criou" class="px-4 py-2 bg-gray-50 border border-gray-300 rounded-lg">
                                    <button id="clearFiltersButton" class="md:col-span-2 w-full h-full bg-gray-200 text-gray-700 font-semibold rounded-lg hover:bg-gray-300">Limpar Filtros</button>
                                </div>
                            </div>
                        </div>

                        <!-- Tabela de Resultados -->
                        <div class="mt-8 overflow-x-auto">
                            <table class="min-w-full divide-y divide-gray-200">
                                <thead class="bg-gray-50">
                                    <tr>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Número</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Data</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Fornecedor</th>
                                        <th class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Total</th>
                                        <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                                        <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Ações</th>
                                    </tr>
                                </thead>
                                <tbody id="searchResultsBody" class="bg-white divide-y divide-gray-200">
                                    <!-- Linhas de resultados serão inseridas aqui -->
                                </tbody>
                            </table>
                        </div>
                        
                        <!-- Botão Voltar -->
                        <div class="mt-6 text-center">
                            <button id="backToMenuFromSearchButton" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">Voltar ao Início</button>
                        </div>
                    </div>
                    
                    <!-- Modal para Simulação de Impressão -->
                    <div id="print-modal" class="fixed inset-0 bg-black bg-opacity-50 hidden items-center justify-center p-4">
                        <div id="print-content" class="bg-white p-8 rounded-lg max-w-2xl w-full">
                            <!-- Conteúdo da impressão simulada aqui -->
                            <h3 class="text-xl font-bold">Pedido #<span id="print-pedido-id"></span></h3>
                            <p>Conteúdo do pedido...</p>
                            <button id="close-print-modal" class="mt-4 w-full py-2 bg-gray-500 text-white rounded-lg">Fechar</button>
                        </div>
                    </div>
                `,
                Dashboard: `
                    <div class="space-y-8">
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                            <div class="bg-white p-6 rounded-xl shadow-lg flex items-center space-x-4">
                                <div class="bg-blue-100 p-4 rounded-full"><i class="fas fa-dollar-sign fa-2x text-blue-600"></i></div>
                                <div>
                                    <p class="text-gray-500 text-sm">Valor Total Gasto</p>
                                    <p id="totalValueCard" class="text-2xl font-bold text-gray-800">R$ 0,00</p>
                                </div>
                            </div>
                            <div class="bg-white p-6 rounded-xl shadow-lg flex items-center space-x-4">
                                <div class="bg-green-100 p-4 rounded-full"><i class="fas fa-receipt fa-2x text-green-600"></i></div>
                                <div>
                                    <p class="text-gray-500 text-sm">Total de Pedidos</p>
                                    <p id="totalOrdersCard" class="text-2xl font-bold text-gray-800">0</p>
                                </div>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                            <div class="bg-white p-6 rounded-xl shadow-lg">
                                <h3 class="text-lg font-semibold text-gray-800 mb-4">Pedidos por Mês</h3>
                                <div class="h-64"><canvas id="monthlyOrdersChart"></canvas></div>
                            </div>
                            <div class="bg-white p-6 rounded-xl shadow-lg">
                                <h3 class="text-lg font-semibold text-gray-800 mb-4">Top 5 Fornecedores (por valor)</h3>
                                <div class="h-64"><canvas id="topSuppliersChart"></canvas></div>
                            </div>
                        </div>

                        <div class="bg-white p-6 rounded-xl shadow-lg">
                             <h3 class="text-lg font-semibold text-gray-800 mb-4">Sugestões de Itens</h3>
                             <div id="aiSuggestions" class="text-gray-700 prose prose-sm max-w-none">
                                <p class="text-gray-500 italic">Nenhuma sugestão disponível.</p>
                             </div>
                        </div>
                        
                        <div class="mt-6 text-center">
                            <button id="backToMenuFromDashboard" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">Voltar ao Início</button>
                        </div>
                    </div>
                `,
                Relatorios: `
                    <div class="max-w-3xl mx-auto bg-white rounded-2xl shadow-xl p-8 space-y-6">
                        <h2 class="text-2xl font-bold text-center text-gray-800">Relatórios de Compras</h2>
                        <form id="reportForm" autocomplete="off" class="space-y-6">
                            <div class="p-4 border border-gray-200 rounded-lg">
                                <h3 class="text-lg font-medium text-gray-900 mb-4">Filtros</h3>
                                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                                    <div>
                                        <label for="startDate" class="block text-sm font-medium text-gray-700">Data Inicial:</label>
                                        <input type="date" id="startDate" name="startDate" class="mt-1 block w-full px-3 py-2 bg-gray-50 border-gray-300 rounded-md">
                                    </div>
                                    <div>
                                        <label for="endDate" class="block text-sm font-medium text-gray-700">Data Final:</label>
                                        <input type="date" id="endDate" name="endDate" class="mt-1 block w-full px-3 py-2 bg-gray-50 border-gray-300 rounded-md">
                                    </div>
                                    <div class="md:col-span-2">
                                        <label for="supplierSelect" class="block text-sm font-medium text-gray-700">Fornecedor:</label>
                                        <select id="supplierSelect" name="supplierSelect" class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 rounded-md">
                                            <option value="todos">Todos os Fornecedores</option>
                                        </select>
                                    </div>
                                </div>
                            </div>

                            <div class="p-4 border border-gray-200 rounded-lg">
                                <h3 class="text-lg font-medium text-gray-900 mb-4">Tipo de Relatório</h3>
                                <div class="flex flex-col space-y-2">
                                    <label class="flex items-center">
                                        <input type="radio" id="reportTypeDetailed" name="reportType" value="detailed" checked class="h-4 w-4 text-blue-600 border-gray-300">
                                        <span class="ml-3 text-sm text-gray-700">Detalhado (por Data/Fornecedor)</span>
                                    </label>
                                    <label class="flex items-center">
                                        <input type="radio" id="reportTypeFinancial" name="reportType" value="financial" class="h-4 w-4 text-blue-600 border-gray-300">
                                        <span class="ml-3 text-sm text-gray-700">Financeiro (Resumo)</span>
                                    </label>
                                </div>
                            </div>
                            
                            <p id="reportMessage" class="text-center text-sm font-medium min-h-[20px]"></p>

                            <div id="reportPreviewArea" class="hidden text-center p-4 bg-blue-50 rounded-lg">
                                <a id="downloadLink" href="#" target="_blank" class="text-blue-600 font-semibold hover:underline">
                                    <i class="fas fa-download mr-2"></i>Baixar Relatório
                                </a>
                            </div>

                            <div class="flex justify-end space-x-4 pt-4 border-t border-gray-200">
                                <button type="button" id="backToMenuFromReportsButton" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">Voltar</button>
                                <button type="button" id="generatePdfButton" class="px-6 py-2 bg-red-600 text-white font-semibold rounded-lg shadow-md hover:bg-red-700"><i class="fas fa-file-pdf mr-2"></i>Gerar PDF</button>
                                <button type="button" id="generateXlsButton" class="px-6 py-2 bg-green-600 text-white font-semibold rounded-lg shadow-md hover:bg-green-700"><i class="fas fa-file-excel mr-2"></i>Gerar XLSX</button>
                            </div>
                        </form>
                    </div>`,
                GerenciarFornecedores: `
                    <div class="bg-white rounded-2xl shadow-xl p-6 sm:p-8">
                        <h2 class="text-2xl font-bold text-center text-gray-800 mb-6">Gerenciar Fornecedores</h2>
                        <p id="fornecedorListMessage" class="text-center text-sm font-medium min-h-[20px] mb-4"></p>
                        <div class="flex justify-start mb-4">
                            <button id="addNewFornecedorBtn" class="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700">
                                <i class="fas fa-plus mr-2"></i>Adicionar Novo
                            </button>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="min-w-full divide-y divide-gray-200">
                                <thead class="bg-gray-50">
                                    <tr>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Código</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Razão Social</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Nome Fantasia</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">CNPJ</th>
                                        <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Grupo</th>
                                        <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Ações</th>
                                    </tr>
                                </thead>
                                <tbody id="fornecedoresTableBody" class="bg-white divide-y divide-gray-200"></tbody>
                            </table>
                        </div>
                        <div class="mt-6 text-center">
                            <button id="backToMenuFromSuppliersButton" class="px-6 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">Voltar ao Início</button>
                        </div>
                    </div>`,
                TrocarSenha: `
                    <div class="max-w-md mx-auto bg-white rounded-2xl shadow-xl p-8 space-y-6">
                        <h2 class="text-2xl font-bold text-center text-gray-800">Trocar Senha</h2>
                        <p id="msgTrocaSenha" class="text-center text-sm font-medium min-h-[20px]"></p>
                        <div class="space-y-4">
                            <div>
                                <label for="senhaAtual" class="block text-sm font-medium text-gray-700">Senha Atual:</label>
                                <input type="password" id="senhaAtual" autocomplete="current-password" class="mt-1 w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg">
                            </div>
                            <div>
                                <label for="novaSenha" class="block text-sm font-medium text-gray-700">Nova Senha:</label>
                                <input type="password" id="novaSenha" autocomplete="new-password" class="mt-1 w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg">
                            </div>
                            <div>
                                <label for="confirmaSenha" class="block text-sm font-medium text-gray-700">Confirme a Nova Senha:</label>
                                <input type="password" id="confirmaSenha" autocomplete="new-password" class="mt-1 w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-lg">
                            </div>
                        </div>
                        <div class="flex space-x-4 pt-4 border-t border-gray-200">
                            <button id="btnVoltarMenuSenha" class="w-full py-3 px-4 bg-gray-500 text-white font-semibold rounded-lg hover:bg-gray-600">Voltar</button>
                            <button id="btnSalvarSenha" class="w-full py-3 px-4 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700">Salvar</button>
                        </div>
                    </div>`,
                GerenciarRascunhos: `
                    <div class="max-w-7xl mx-auto">
                        <header class="mb-6">
                            <div class="flex items-center justify-between">
                                <div>
                                    <h1 class="text-3xl font-bold text-gray-800">Gerenciar Rascunhos</h1>
                                    <p class="text-gray-500">Visualize, edite ou finalize seus rascunhos salvos.</p>
                                </div>
                                <button id="backToMenuFromRascunhosButton" class="px-4 py-2 bg-gray-600 text-white font-semibold rounded-lg shadow-md hover:bg-gray-700">
                                    <i class="fas fa-arrow-left mr-2"></i>Voltar ao Menu
                                </button>
                            </div>
                            <p id="rascunhosMessage" class="text-center text-sm font-medium min-h-[20px] mt-2"></p>
                        </header>

                        <main class="bg-white p-6 rounded-xl shadow-lg">
                            <div class="mb-4 flex items-center justify-between">
                                <h3 class="text-lg font-bold text-gray-700">Rascunhos Salvos</h3>
                                <span id="totalRascunhos" class="text-sm text-gray-500">Carregando...</span>
                            </div>
                            
                            <!-- Filtros -->
                            <div class="mb-6 grid grid-cols-1 md:grid-cols-3 gap-4">
                                <div>
                                    <label for="filtroFornecedor" class="block text-sm font-medium text-gray-700">Fornecedor</label>
                                    <select id="filtroFornecedor" class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 rounded-md">
                                        <option value="">Todos os fornecedores</option>
                                    </select>
                                </div>
                                <div>
                                    <label for="filtroDataInicio" class="block text-sm font-medium text-gray-700">Data Início</label>
                                    <input type="date" id="filtroDataInicio" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md">
                                </div>
                                <div>
                                    <label for="filtroDataFim" class="block text-sm font-medium text-gray-700">Data Fim</label>
                                    <input type="date" id="filtroDataFim" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md">
                                </div>
                            </div>

                            <!-- Tabela de Rascunhos -->
                            <div class="overflow-x-auto">
                                <table class="min-w-full divide-y divide-gray-200">
                                    <thead class="bg-gray-50">
                                        <tr>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">ID</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Data</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Fornecedor</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Itens</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Última Edição</th>
                                            <th class="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Ações</th>
                                        </tr>
                                    </thead>
                                    <tbody id="rascunhosTableBody" class="bg-white divide-y divide-gray-200">
                                        <!-- Conteúdo será carregado dinamicamente -->
                                    </tbody>
                                </table>
                            </div>

                            <!-- Estado vazio -->
                            <div id="emptyStateRascunhos" class="hidden text-center py-10">
                                <div class="bg-gray-100 rounded-full h-16 w-16 flex items-center justify-center mx-auto mb-4">
                                    <i class="fas fa-inbox text-gray-400 text-xl"></i>
                                </div>
                                <p class="text-gray-500 text-lg mb-2">Nenhum rascunho encontrado</p>
                                <p class="text-gray-400 text-sm">Comece criando um novo pedido e salvando como rascunho.</p>
                                <button onclick="criarNovoPedidoFromRascunho()" class="mt-4 px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-lg transition-colors duration-200">
                                    <i class="fas fa-plus mr-2"></i>Criar Novo Pedido
                                </button>
                            </div>
                        </main>
                    </div>`
            };
            return templates[pageName] || null;
        }

        // ===============================================
        // FUNÇÃO AUXILIAR PARA CRIAR NOVO PEDIDO
        // ===============================================
        function criarNovoPedidoFromRascunho() {
            // Limpar dados de visualização
            localStorage.removeItem('modoVisualizacao');
            localStorage.removeItem('origemVisualizacao'); 
            localStorage.removeItem('pedidoParaVisualizar');
            localStorage.removeItem('empresaOriginalAdmin');
            localStorage.removeItem('numeroPedidoVisualizacao');
            
            // Carregar tela de novo pedido
            loadPageContent('Pedido', setupPedidoScreen);
        }

        // ===============================================
        // FUNÇÃO PARA SALVAR EDIÇÃO DE PEDIDO
        // ===============================================
        
        function salvarEdicaoPedido() {
            console.log('💾 Iniciando salvamento de edição do pedido');
            
            // Verificar se ainda está em modo de edição válido
            if (!EdicaoPedido.validarPermissaoEdicao()) {
                showToast('Tempo de edição expirado ou dados inválidos. Redirecionando...', 'warning', 'Tempo Expirado');
                setTimeout(() => {
                    EdicaoPedido.limpar();
                    loadPageContent('BuscarPedido', setupBuscaScreen);
                }, 2000);
                return;
            }
            
            const salvarButton = document.getElementById('salvarPedidoButton');
            const dadosOriginais = EdicaoPedido.getDados();
            
            if (!dadosOriginais) {
                showToast('Erro: Dados do pedido não encontrados.', 'error', 'Erro Interno');
                return;
            }
            
            // Coletar dados do formulário (reutilizando a lógica existente)
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            const idDaEmpresa = empresaAtual.id || empresaAtual.codigo;
            
            // Validar campos obrigatórios básicos
            const fornecedor = document.getElementById('fornecedorPedido')?.value;
            if (!fornecedor) {
                showToast('Por favor, selecione um fornecedor.', 'error', 'Campo Obrigatório');
                return;
            }
            
            // Coletar itens
            const itens = [];
            const itensContainer = document.getElementById('itensContainer');
            if (!itensContainer || itensContainer.children.length === 0) {
                showToast('Adicione pelo menos um item ao pedido.', 'error', 'Itens Obrigatórios');
                return;
            }
            
            let totalGeral = 0;
            let itemValido = true;
            
            Array.from(itensContainer.children).forEach((itemDiv, index) => {
                const descricao = document.getElementById(`itemDescricao${index}`)?.value || '';
                const unidade = document.getElementById(`itemUnidade${index}`)?.value || '';
                const quantidade = parseFloat(document.getElementById(`itemQuantidade${index}`)?.value || '0');
                const precoUnitario = parseFloat(document.getElementById(`itemPrecoUnitario${index}`)?.value || '0');
                
                if (!descricao.trim()) {
                    showToast(`Descrição do item ${index + 1} é obrigatória.`, 'error', 'Item Incompleto');
                    itemValido = false;
                    return;
                }
                
                if (quantidade <= 0 || precoUnitario <= 0) {
                    showToast(`Quantidade e preço do item ${index + 1} devem ser maiores que zero.`, 'error', 'Valores Inválidos');
                    itemValido = false;
                    return;
                }
                
                const totalItem = quantidade * precoUnitario;
                totalGeral += totalItem;
                
                itens.push({
                    descricao: descricao.trim(),
                    unidade: unidade.trim(),
                    quantidade: quantidade,
                    precoUnitario: precoUnitario,
                    totalItem: totalItem
                });
            });
            
            if (!itemValido) {
                return;
            }
            
            // Preparar dados para edição
            const dadosEdicao = {
                numeroDoPedido: dadosOriginais.numeroDoPedido,
                fornecedor: fornecedor,
                nomeVeiculo: document.getElementById('nomeVeiculo')?.value || '',
                placaVeiculo: document.getElementById('placaVeiculo')?.value || '',
                observacoes: document.getElementById('observacoes')?.value || '',
                itens: itens,
                totalGeral: totalGeral,
                empresaId: idDaEmpresa
            };
            
            console.log('📤 Dados para edição:', dadosEdicao);
            
            // Desabilitar botão e mostrar loading
            if (salvarButton) {
                salvarButton.disabled = true;
                salvarButton.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Salvando Alterações...';
            }
            
            showToast('Salvando alterações no pedido...', 'info', 'Processando');
            
            // Enviar para o backend
            google.script.run
                .withSuccessHandler(response => {
                    console.log('✅ Resposta da edição:', response);
                    
                    if (salvarButton) {
                        salvarButton.disabled = false;
                        salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Alterações';
                    }
                    
                    if (response.status === 'success') {
                        showToast('Pedido editado com sucesso!', 'success', 'Sucesso');
                        EdicaoPedido.limpar();
                        
                        setTimeout(() => {
                            loadPageContent('BuscarPedido', setupBuscaScreen);
                        }, 2000);
                    } else {
                        showToast(response.message || 'Erro ao salvar alterações.', 'error', 'Erro no Salvamento');
                    }
                })
                .withFailureHandler(error => {
                    console.error('❌ Erro ao salvar edição:', error);
                    
                    if (salvarButton) {
                        salvarButton.disabled = false;
                        salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Alterações';
                    }
                    
                    showToast('Erro de comunicação ao salvar alterações.', 'error', 'Erro de Comunicação');
                })
                .editarPedido(dadosEdicao);
        }

        // ===============================================
        // FUNÇÕES UTILITÁRIAS GLOBAIS
        // ===============================================
        
        function showGlobalModal(title, message, type = 'info') {
            const modal = document.getElementById('global-modal');
            if(!modal) return;
            const modalTitle = modal.querySelector('#modal-title');
            const modalMessage = modal.querySelector('#modal-message');
            if(modalTitle) modalTitle.textContent = title;
            if(modalMessage) modalMessage.textContent = message;
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        document.getElementById('modal-close-btn').addEventListener('click', () => {
            const modal = document.getElementById('global-modal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        });
        document.getElementById('global-modal').addEventListener('click', (event) => {
            if (event.target === document.getElementById('global-modal')) {
                const modal = document.getElementById('global-modal');
                modal.classList.add('hidden');
                modal.classList.remove('flex');
            }
        });

        function showMessage(text, type = 'error', targetId) {
            // Função de compatibilidade - agora usa o sistema de toasts
            // Se targetId for fornecido, tenta usar toast com fallback
            if (targetId) {
                return showToast(text, type, null, 5000, targetId);
            }
            
            // Sem targetId, usa toast diretamente
            return showToast(text, type);
        }

        
        // ===============================================
        // SISTEMA DE TOASTS - NOTIFICAÇÕES MODERNAS
        // ===============================================
        
        let toastCounter = 0;
        
        function showToast(message, type = 'info', title = null, duration = 5000, targetId = null) {
            // Gerar ID único para o toast
            toastCounter++;
            const toastId = `toast-${toastCounter}`;
            
            // Se targetId for fornecido, verificar se elemento existe e está visível
            if (targetId) {
                const targetElement = document.getElementById(targetId);
                if (targetElement && targetElement.offsetParent !== null) {
                    // Elemento existe e está visível - usar sistema antigo
                    if (targetElement) {
                        targetElement.textContent = message;
                        targetElement.classList.remove('text-red-500', 'text-green-500', 'text-blue-500', 'text-yellow-500');
                        if (type === 'error') targetElement.classList.add('text-red-500');
                        if (type === 'success') targetElement.classList.add('text-green-500');
                        if (type === 'info') targetElement.classList.add('text-blue-500');
                        if (type === 'warning') targetElement.classList.add('text-yellow-500');
                        return toastId; // Retorna ID mesmo quando usa sistema antigo
                    }
                }
                // Se elemento não existe ou não está visível, continua para usar toast
            }
            
            // Configurações por tipo
            const toastConfig = {
                success: {
                    icon: 'fa-check-circle',
                    title: title || 'Sucesso',
                    class: 'toast-success'
                },
                error: {
                    icon: 'fa-exclamation-circle',
                    title: title || 'Erro',
                    class: 'toast-error'
                },
                warning: {
                    icon: 'fa-exclamation-triangle',
                    title: title || 'Atenção',
                    class: 'toast-warning'
                },
                info: {
                    icon: 'fa-info-circle',
                    title: title || 'Informação',
                    class: 'toast-info'
                }
            };
            
            const config = toastConfig[type] || toastConfig.info;
            
            // Criar elemento do toast
            const toastElement = document.createElement('div');
            toastElement.id = toastId;
            toastElement.className = `toast ${config.class}`;
            
            toastElement.innerHTML = `
                <div class="toast-header">
                    <div class="toast-title">
                        <i class="fas ${config.icon} toast-icon"></i>
                        ${config.title}
                    </div>
                    <button class="toast-close" onclick="hideToast('${toastId}')">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                <div class="toast-message">${message}</div>
                ${duration > 0 ? '<div class="toast-progress"></div>' : ''}
            `;
            
            // Adicionar ao container
            const container = document.getElementById('toast-container');
            if (container) {
                container.appendChild(toastElement);
                
                // Trigger animation
                setTimeout(() => {
                    toastElement.classList.add('show');
                }, 10);
                
                // Auto-remove se duration for especificado
                if (duration > 0) {
                    setTimeout(() => {
                        hideToast(toastId);
                    }, duration);
                }
            }
            
            return toastId;
        }
        
        
        function hideToast(toastId) {
            const toast = document.getElementById(toastId);
            if (toast) {
                toast.classList.remove('show');
                toast.classList.add('hide');
                
                setTimeout(() => {
                    if (toast.parentNode) {
                        toast.parentNode.removeChild(toast);
                    }
                }, 300);
            }
        }
        
        function hideAllToasts() {
            const container = document.getElementById('toast-container');
            if (container) {
                const toasts = container.querySelectorAll('.toast');
                toasts.forEach(toast => {
                    hideToast(toast.id);
                });
            }
        }
        
        // Função de conveniência para substituir showMessage
        function showToastMessage(text, type = 'info', targetId = null) {
            // Se targetId for fornecido, usar o método antigo para compatibilidade
            if (targetId) {
                return showMessage(text, type, targetId);
            }
            
            // Usar o novo sistema de toasts
            return showToast(text, type);
        }
        
        // ===============================================
        // LÓGICA DA TELA DE LOGIN
        // ===============================================
        function setupLoginScreen() {
            const loginButton = document.getElementById('loginButton');
            const createAccountButton = document.getElementById('createAccountButton');
            const usernameField = document.getElementById('username');
            const screenLogin = document.getElementById('screen-login');
            
            if (loginButton) {
                loginButton.addEventListener('click', logar);
            }
            
            if (createAccountButton) {
                createAccountButton.addEventListener('click', () => showScreen('screen-cadastro'));
            }
            
            if (usernameField) {
                usernameField.addEventListener('blur', carregarEmpresasParaLogin);
            }
            
            if (screenLogin) {
                screenLogin.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') logar();
                });
            }
        }

        function carregarEmpresasParaLogin() {
            const username = document.getElementById('username').value.trim();
            const selectEmpresa = document.getElementById('selectEmpresaLogin');
            
            if (!username) {
                selectEmpresa.innerHTML = `<option value="">Primeiro, digite o usuário</option>`;
                return;
            }
            
            selectEmpresa.innerHTML = `<option value="">Carregando empresas...</option>`;
            
            google.script.run
                .withSuccessHandler(response => {
                    if (!selectEmpresa || !response || !response.empresas) {
                        if (selectEmpresa) {
                            selectEmpresa.innerHTML = `<option value="">Nenhuma empresa disponível</option>`;
                        }
                        return;
                    }
                    
                    selectEmpresa.innerHTML = `<option value="">Selecione a empresa</option>`;
                    
                    response.empresas.forEach(emp => {
                        const option = document.createElement('option');
                        option.value = emp.id;
                        option.textContent = emp.nome;
                        selectEmpresa.appendChild(option);
                    });
                    
                    if (response.defaultEmpresaId) {
                        selectEmpresa.value = response.defaultEmpresaId;
                    } else if (selectEmpresa.options.length > 1) {
                        selectEmpresa.selectedIndex = 1;
                    }
                })
                .withFailureHandler(err => {
                    console.error('Erro ao obter empresas:', err);
                    if (selectEmpresa) {
                        selectEmpresa.innerHTML = `<option value="">Erro ao carregar</option>`;
                    }
                })
                .obterEmpresasDoUsuario(username);
        }

        function logar() {
            const username = document.getElementById('username').value.trim();
            const password = document.getElementById('password').value.trim();
            const selectEmpresa = document.getElementById('selectEmpresaLogin');
            const empresaSelecionadaId = selectEmpresa.value;
            const empresaSelecionadaNome = selectEmpresa.options[selectEmpresa.selectedIndex].text;
            const loginButton = document.getElementById('loginButton');

            if (!username || !password || !empresaSelecionadaId) {
                showToast('Preencha usuário, senha e selecione a empresa.', 'warning', 'Campos Obrigatórios');
                return;
            }

            loginButton.disabled = true;
            loginButton.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Entrando...';
            
            google.script.run
                .withSuccessHandler(res => {
                    if (res.status === 'ok' && res.empresa) {
                        localStorage.setItem('nome', res.nome);
                        localStorage.setItem('perfil', res.perfil);
                        localStorage.setItem('usuarioLogado', username);
                        localStorage.setItem('empresaSelecionada', JSON.stringify(res.empresa));
                        
                        showToast(`Bem-vindo, ${res.nome}!`, 'success', 'Login Realizado');
                        
                        google.script.run
                            .withSuccessHandler(() => {
                                showScreen('screen-main');
                                setupMainScreen();
                            })
                            .registrarLoginUsuario(res.idUsuario, res.nomeUsuario, empresaSelecionadaId, empresaSelecionadaNome);
                    } else {
                        showToast(res.message || 'Usuário, senha ou empresa inválidos!', 'error', 'Falha no Login');
                        loginButton.disabled = false;
                        loginButton.textContent = 'Entrar';
                    }
                })
                .withFailureHandler(err => {
                    showToast('Erro de comunicação. Tente novamente.', 'error', 'Falha na Conexão');
                    loginButton.disabled = false;
                    loginButton.textContent = 'Entrar';
                    console.error("Erro em validarLogin:", err);
                })
                .validarLogin(username, password, empresaSelecionadaId);
        }

        // ===============================================
        // LÓGICA DA TELA DE CADASTRO DE USUÁRIO
        // ===============================================
        function setupCadastroScreen() {
            const registerButton = document.getElementById('registerButton');
            const backButton = document.getElementById('backToLoginButton');
            const nomeCompletoInput = document.getElementById('name');
            const usernameInput = document.getElementById('username-cadastro');
            
            if (registerButton) {
                registerButton.addEventListener('click', cadastrarUsuario);
            }
            
            if (backButton) {
                backButton.addEventListener('click', () => showScreen('screen-login'));
            }
            
            if (nomeCompletoInput && usernameInput) {
                nomeCompletoInput.addEventListener('input', function() {
                    const nomeDigitado = this.value;
                    const nomes = nomeDigitado.trim().toLowerCase().split(' ');
                    const primeiro = nomes[0];
                    const ultimo = nomes.length > 1 ? nomes[nomes.length - 1] : '';
                    let usernameSugerido = ultimo ? `${primeiro}.${ultimo}` : primeiro;
                    usernameSugerido = usernameSugerido.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9.]/g, '');
                    usernameInput.value = usernameSugerido;
                });
            }
        }

        function cadastrarUsuario() {
            const nome = document.getElementById('name').value.trim();
            const usuario = document.getElementById('username-cadastro').value.trim();
            const senha = document.getElementById('password-cadastro').value.trim();
            const registerButton = document.getElementById('registerButton');

            if (!nome || !usuario || !senha) {
                showToast('Preencha todos os campos!', 'error', 'message-cadastro');
                return;
            }
            if (senha.length < 6) {
                showToast('A senha deve ter no mínimo 6 caracteres.', 'error', 'message-cadastro');
                return;
            }

            registerButton.disabled = true;
            registerButton.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Cadastrando...';

            google.script.run
                .withSuccessHandler(res => {
                    if (res.status === 'ok') {
                        showGlobalModal('Sucesso!', res.message || 'Solicitação enviada. Aguarde aprovação.');
                        showScreen('screen-login');
                    } else {
                        showToast(res.message, 'error', 'message-cadastro');
                    }
                    registerButton.disabled = false;
                    registerButton.textContent = 'Solicitar Cadastro';
                })
                .withFailureHandler(err => {
                    showToast('Erro de comunicação.', 'error', 'message-cadastro');
                    registerButton.disabled = false;
                    registerButton.textContent = 'Solicitar Cadastro';
                    console.error("Erro em criarUsuario:", err);
                })
                .criarUsuario(nome, usuario, senha, '', 'usuario');
        }

        // ===============================================
        // LÓGICA DA TELA PRINCIPAL E MENU
        // ===============================================
        function setupMainScreen() {
            const nome = localStorage.getItem("nome") || "Usuário";
            const perfil = (localStorage.getItem("perfil") || "").toLowerCase();
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');

            document.getElementById('nome').textContent = nome;
            document.getElementById('perfil').textContent = perfil;
            document.getElementById('empresaInfo').textContent = empresaAtual.nome || empresaAtual.empresa || "N/A";
            document.getElementById('logoutButton').addEventListener('click', logout);
            
            loadPageContent('Menu', setupMenuScreen);
        }
        
        function logout() {
            // Limpar todos os dados do localStorage
            localStorage.removeItem('usuarioLogado');
            localStorage.removeItem('nome');
            localStorage.removeItem('perfil');
            localStorage.removeItem('empresaSelecionada');
            localStorage.removeItem('accessToken');
            localStorage.removeItem('empresaOriginalAdmin');
            localStorage.removeItem('pedidoParaVisualizar');
            localStorage.removeItem('modoVisualizacao');
            
            // Redirecionar para a tela de login
            showScreen('screen-login');
            setupLoginScreen();
        }
        
        function setupMenuScreen() {
            const perfil = (localStorage.getItem("perfil") || "").toLowerCase();
            const usuarioLogado = localStorage.getItem('usuarioLogado');
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            
            const switcherSelect = document.getElementById('selectEmpresa');
            google.script.run
                .withSuccessHandler(response => {
                    if (!switcherSelect || !response || !Array.isArray(response.empresas)) return;
                    switcherSelect.innerHTML = '';
                    response.empresas.forEach(emp => {
                        const option = document.createElement('option');
                        option.value = emp.id;
                        option.textContent = emp.nome;
                        switcherSelect.appendChild(option);
                    });
                    if (empresaAtual.id) {
                        switcherSelect.value = empresaAtual.id;
                    }
                })
                .obterEmpresasDoUsuario(usuarioLogado);
            
            document.getElementById('btnSelecionarEmpresa').addEventListener('click', () => {
                const novaEmpresaId = switcherSelect.value;
                showGlobalModal('Atenção', 'Trocando de empresa...', 'info');
                google.script.run
                    .withSuccessHandler(novaEmpresaObjeto => {
                        if (novaEmpresaObjeto) {
                            localStorage.setItem('empresaSelecionada', JSON.stringify(novaEmpresaObjeto));
                            setupMainScreen();
                            document.getElementById('global-modal').classList.add('hidden');
                        }
                    })
                    ._getEmpresaDataById(novaEmpresaId);
            });

            const menuOptionsGroup = document.getElementById('menuOptionsGroup');
            menuOptionsGroup.innerHTML = '';

            createMenuButton('Novo Pedido', () => {
                // Limpar dados de visualização antes de criar novo pedido
                localStorage.removeItem('modoVisualizacao');
                localStorage.removeItem('origemVisualizacao');
                localStorage.removeItem('pedidoParaVisualizar');
                localStorage.removeItem('empresaOriginalAdmin');
                localStorage.removeItem('numeroPedidoVisualizacao');
                console.log('🔧 Variáveis de visualização limpas para novo pedido');
                loadPageContent('Pedido', setupPedidoScreen);
            }, menuOptionsGroup);
            createMenuButton('Gerenciar Rascunhos', () => loadPageContent('GerenciarRascunhos', setupGerenciarRascunhosScreen), menuOptionsGroup);
            createMenuButton('Cadastro de Fornecedores', () => loadPageContent('CadastroFornecedor', setupCadastroFornecedorScreen), menuOptionsGroup);
            createMenuButton('Buscar Pedido', () => loadPageContent('BuscarPedido', setupBuscaScreen), menuOptionsGroup);
            createMenuButton('Trocar Senha', () => loadPageContent('TrocarSenha', setupTrocarSenhaScreen), menuOptionsGroup);
            
            if (perfil === "admin") {
                createMenuButton('Dashboard Completo', () => loadPageContent('Dashboard', setupDashboardScreen), menuOptionsGroup);
                createMenuButton('Relatórios', () => loadPageContent('Relatorios', setupRelatoriosScreen), menuOptionsGroup);
                createMenuButton('Gerenciar Usuários', () => loadPageContent('GerenciarUsuarios', setupGerenciarUsuariosScreen), menuOptionsGroup);
                createMenuButton('Gerenciar Fornecedores', () => loadPageContent('GerenciarFornecedores', setupGerenciarFornecedoresScreen), menuOptionsGroup);
            }
            
            loadAdminDashboardCards();
            
        function createMenuButton(text, onClick, container) {
            const button = document.createElement('button');
            button.className = 'w-full text-left px-4 py-3 bg-white border border-gray-200 rounded-lg hover:bg-gray-50 hover:border-blue-300 transition-colors duration-200 flex items-center justify-between group';
            button.innerHTML = `
                <span class="text-gray-700 group-hover:text-blue-600 font-medium">${text}</span>
                <i class="fas fa-chevron-right text-gray-400 group-hover:text-blue-600"></i>
            `;
            button.addEventListener('click', onClick);
            container.appendChild(button);
        }

        function loadAdminDashboardCards() {
        const perfil = (localStorage.getItem("perfil") || "").toLowerCase();
        const dashboardContainer = document.getElementById('company-dashboard-container');
        if (!dashboardContainer) return;

        const contentArea = dashboardContainer.querySelector('#dashboard-content-area');
        if (!contentArea) return;

         if (perfil !== 'admin') {
            // Dashboard simplificado para usuários normais - mostrar apenas pedidos da empresa atual
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            const idDaEmpresa = empresaAtual.id || empresaAtual.codigo;
            
            contentArea.innerHTML = `<div class="text-center py-10 text-gray-500"><div class="spinner mx-auto"></div><p class="mt-4">Carregando seus últimos pedidos...</p></div>`;

            google.script.run
                .withSuccessHandler(response => {
                    contentArea.innerHTML = '';
                    
                    if (response.status === 'success' && response.data.length > 0) {
                        // Ordenar pedidos por data decrescente
                        const pedidosOrdenados = response.data.sort((a, b) => {
                            return new Date(b.data + 'T12:00:00') - new Date(a.data + 'T12:00:00');
                        });
                        
                        // Pegar os últimos 5 pedidos
                        const ultimosPedidos = pedidosOrdenados.slice(0, 5);
                        
                        const cardContainer = document.createElement('div');
                        cardContainer.className = 'bg-white p-6 rounded-xl shadow-lg border border-gray-200 mb-6';
                        cardContainer.innerHTML = `
                            <div class="flex items-center mb-4">
                                <div class="bg-gradient-to-r from-blue-500 to-blue-600 text-white rounded-full h-12 w-12 flex items-center justify-center mr-3">
                                    <i class="fas fa-receipt text-lg"></i>
                                </div>
                                <h3 class="text-xl font-bold text-gray-800">Seus Últimos Pedidos - ${empresaAtual.nome}</h3>
                            </div>
                            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4" id="pedidos-usuario-cards"></div>
                        `;
                        
                        contentArea.appendChild(cardContainer);
                        
                        const pedidosContainer = document.getElementById('pedidos-usuario-cards');
                        
                        ultimosPedidos.forEach((pedido, index) => {
                            console.log(`🔍 DEBUG - Pedido ${index + 1}:`, pedido);
                            console.log(`🔍 DEBUG - Estado do pedido ${index + 1}:`, pedido.estadoFornecedor);
                            
                            const valorTotal = `R$ ${parseFloat(pedido.totalGeral || 0).toFixed(2).replace('.', ',')}`;
                            const dataFormatada = new Date(pedido.data + 'T12:00:00').toLocaleDateString('pt-BR');
                            const isRecente = index === 0;
                            
                            const pedidoCard = document.createElement('div');
                            pedidoCard.className = `bg-gradient-to-br ${isRecente ? 'from-green-50 to-green-100 border-green-300' : 'from-gray-50 to-gray-100 border-gray-200'} rounded-lg p-4 border`;
                            pedidoCard.innerHTML = `
                                <div class="flex items-center justify-between mb-3">
                                    <div class="flex items-center space-x-2">
                                        <div class="${isRecente ? 'bg-green-500' : 'bg-blue-500'} text-white rounded-full h-8 w-8 flex items-center justify-center">
                                            <i class="fas ${isRecente ? 'fa-star' : 'fa-file-invoice'} text-sm"></i>
                                        </div>
                                        <span class="text-xs font-medium ${isRecente ? 'text-green-700' : 'text-blue-700'} uppercase tracking-wide">
                                            ${isRecente ? 'Mais Recente' : `#${index + 1}`}
                                        </span>
                                    </div>
                                    <span class="text-lg font-bold ${isRecente ? 'text-green-800' : 'text-gray-800'}">${valorTotal}</span>
                                </div>
                                
                                <div class="space-y-2 text-sm mb-3">
                                    <p class="text-gray-700"><span class="font-semibold">Pedido:</span> #${pedido.numeroDoPedido}</p>
                                    <p class="text-gray-700"><span class="font-semibold">Data:</span> ${dataFormatada}</p>
                                    <p class="text-gray-700"><span class="font-semibold">Fornecedor:</span> ${pedido.fornecedor}</p>
                                    <p class="text-gray-700"><span class="font-semibold">Estado:</span> ${pedido.estadoFornecedor || 'N/A'}</p>
                                </div>
                                
                                <div class="flex gap-2 mt-3 pt-3 border-t border-gray-200">
                                    <button onclick="visualizarPedidoAdmin('${pedido.numeroDoPedido}', '${empresaAtual.id}', '${empresaAtual.nome}')" 
                                            class="flex-1 text-xs px-3 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-md transition-colors duration-200 flex items-center justify-center gap-1">
                                        <i class="fas fa-eye"></i>
                                        <span>Visualizar</span>
                                    </button>
                                    <button onclick="abrirImpressaoPedido('${pedido.numeroDoPedido}')" 
                                            class="flex-1 text-xs px-3 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-md transition-colors duration-200 flex items-center justify-center gap-1">
                                        <i class="fas fa-print"></i>
                                        <span>Imprimir</span>
                                    </button>
                                </div>
                            `;
                            
                            pedidosContainer.appendChild(pedidoCard);
                        });
                    } else {
                        contentArea.innerHTML = `
                            <div class="bg-white p-6 rounded-xl shadow-lg text-center">
                                <div class="bg-gray-100 rounded-full h-16 w-16 flex items-center justify-center mx-auto mb-4">
                                    <i class="fas fa-inbox text-gray-400 text-xl"></i>
                                </div>
                                <p class="text-gray-500 text-lg mb-2">Nenhum pedido encontrado</p>
                                <p class="text-gray-400 text-sm">Você ainda não criou nenhum pedido para esta empresa.</p>
                                <button onclick="criarNovoPedidoFromRascunho()" class="mt-4 px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-lg transition-colors duration-200">
                                    <i class="fas fa-plus mr-2"></i>Criar Primeiro Pedido
                                </button>
                            </div>
                        `;
                    }
                })
                .withFailureHandler(err => {
                     contentArea.innerHTML = `
                        <div class="bg-white p-6 rounded-xl shadow-lg text-center">
                            <div class="bg-red-100 rounded-full h-16 w-16 flex items-center justify-center mx-auto mb-4">
                                <i class="fas fa-exclamation-triangle text-red-400 text-xl"></i>
                            </div>
                            <p class="text-red-500 text-lg mb-2">Erro ao carregar pedidos</p>
                            <p class="text-gray-500 text-sm">Tente novamente em alguns instantes.</p>
                        </div>
                     `;
                     console.error("Erro ao buscar pedidos do usuário:", err);
                })
                .buscarPedidos("", idDaEmpresa);
            return;
        }
        
        // Lógica exclusiva para administradores
        const usuarioLogado = localStorage.getItem('usuarioLogado');
        contentArea.innerHTML = `<div class="text-center py-10 text-gray-500"><div class="spinner mx-auto"></div><p class="mt-4">Carregando resumos das empresas...</p></div>`;
        
        google.script.run
            .withSuccessHandler(function(dadosCompletos) {
                console.log("Resposta da API:", dadosCompletos);
                
                contentArea.innerHTML = '';
                
                // Verificar se temos dados válidos
                if (!dadosCompletos || !dadosCompletos.empresas) {
                    contentArea.innerHTML = '<div class="bg-white p-6 rounded-xl shadow-lg text-center"><p class="text-red-500 italic">Dados inválidos retornados.</p></div>';
                    return;
                }
                
                var arrayEmpresas = dadosCompletos.empresas;
                
                if (arrayEmpresas.length > 0) {
                    console.log("Processando", arrayEmpresas.length, "empresas");
                    
                    arrayEmpresas.forEach(function(empresaAtual) {
                        console.log("Criando card para:", empresaAtual.nome);
                        
                        var cardEmpresa = document.createElement('div');
                        cardEmpresa.className = 'bg-white p-6 rounded-xl shadow-lg border border-gray-200 hover:shadow-xl transition-shadow duration-300 mb-6';
                        cardEmpresa.innerHTML = `
                            <div class="flex items-center mb-4">
                                <div class="bg-gradient-to-r from-blue-500 to-blue-600 text-white rounded-full h-12 w-12 flex items-center justify-center mr-3">
                                    <i class="fas fa-building text-lg"></i>
                                </div>
                                <h3 class="text-xl font-bold text-gray-800">${empresaAtual.nome}</h3>
                            </div>
                            <div id="card-content-${empresaAtual.id}" class="text-center">
                                <div class="flex items-center justify-center space-x-2 text-gray-500">
                                    <div class="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-500"></div>
                                    <span class="text-sm">Carregando últimos pedidos...</span>
                                </div>
                            </div>
                        `;
                        contentArea.appendChild(cardEmpresa);
                        
                        // Buscar os últimos 2 pedidos desta empresa
                        google.script.run
                            .withSuccessHandler(function(responsePedidos) {
                                var elementoCard = document.getElementById(`card-content-${empresaAtual.id}`);
                                
                                if (responsePedidos.status === 'success' && responsePedidos.data.length > 0) {
                                    // Ordenar por data decrescente para garantir que são os mais recentes
                                    var pedidosOrdenados = responsePedidos.data.sort((a, b) => {
                                        return new Date(b.data + 'T12:00:00') - new Date(a.data + 'T12:00:00');
                                    });
                                    
                                    // Pegar os 2 mais recentes
                                    var ultimosDoisPedidos = pedidosOrdenados.slice(0, 2);
                                    
                                    var cardsHtml = '<div class="grid grid-cols-1 md:grid-cols-2 gap-4">';
                                    
                                    ultimosDoisPedidos.forEach(function(pedido, index) {
                                        var valorTotal = `R$ ${parseFloat(pedido.totalGeral || 0).toFixed(2).replace('.', ',')}`;
                                        var dataUltimoPedido = new Date(pedido.data + 'T12:00:00').toLocaleDateString('pt-BR');
                                        var isRecente = index === 0; // O primeiro é o mais recente
                                        
                                        cardsHtml += `
                                            <div class="bg-gradient-to-br from-gray-50 to-gray-100 rounded-lg p-4 border ${isRecente ? 'border-green-300 bg-gradient-to-br from-green-50 to-green-100' : 'border-gray-200'}">
                                                <div class="flex items-center justify-between mb-3">
                                                    <div class="flex items-center space-x-2">
                                                        <div class="${isRecente ? 'bg-green-500' : 'bg-blue-500'} text-white rounded-full h-8 w-8 flex items-center justify-center">
                                                            <i class="fas ${isRecente ? 'fa-star' : 'fa-receipt'} text-sm"></i>
                                                        </div>
                                                        <span class="text-xs font-medium ${isRecente ? 'text-green-700' : 'text-blue-700'} uppercase tracking-wide">
                                                            ${isRecente ? 'Mais Recente' : 'Anterior'}
                                                        </span>
                                                    </div>
                                                    <span class="text-lg font-bold ${isRecente ? 'text-green-800' : 'text-gray-800'}">${valorTotal}</span>
                                                </div>
                                                
                                                <div class="space-y-2 text-sm mb-3">
                                                    <p class="text-gray-700"><span class="font-semibold">Pedido:</span> #${pedido.numeroDoPedido}</p>
                                                    <!-- Corrigir para considerar data e hora do pedido -->
                                                    <p class="text-gray-700"><span class="font-semibold">Data:</span> ${
                                                        (() => {
                                                            let dataFormatada;
                                                            if (pedido.data && pedido.data.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
                                                                dataFormatada = new Date(pedido.data.replace(' ', 'T')).toLocaleString('pt-BR');
                                                            } else if (pedido.data) {
                                                                dataFormatada = new Date(pedido.data).toLocaleString('pt-BR');
                                                            } else {
                                                                dataFormatada = 'N/A';
                                                            }
                                                            return dataFormatada;
                                                        })()
                                                    }</p>
                                                    <p class="text-gray-700"><span class="font-semibold">Fornecedor:</span> ${pedido.fornecedor}</p>
                                                    <p class="text-gray-700"><span class="font-semibold">Estado:</span> ${pedido.estadoFornecedor || 'N/A'}</p>
                                                </div>
                                                
                                                <div class="flex gap-2 mt-3 pt-3 border-t border-gray-200">
                                                    <button onclick="visualizarPedidoAdmin('${pedido.numeroDoPedido}', '${empresaAtual.id}', '${empresaAtual.nome}')" 
                                                            class="flex-1 text-xs px-3 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-md transition-colors duration-200 flex items-center justify-center gap-1">
                                                        <i class="fas fa-eye"></i>
                                                        <span>Visualizar</span>
                                                    </button>
                                                    <button onclick="abrirImpressaoPedidoAdmin('${pedido.numeroDoPedido}', '${empresaAtual.id}', '${empresaAtual.nome}')" 
                                                            class="flex-1 text-xs px-3 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-md transition-colors duration-200 flex items-center justify-center gap-1">
                                                        <i class="fas fa-print"></i>
                                                        <span>Imprimir</span>
                                                    </button>
                                                </div>
                                            </div>
                                        `;
                                   
                                   
                                    });
                                    
                                    cardsHtml += '</div>';
                                    
                                    // Se houver apenas 1 pedido, mostrar mensagem
                                    if (ultimosDoisPedidos.length === 1) {
                                        cardsHtml += '<p class="text-center text-gray-500 text-sm mt-4 italic">Esta empresa possui apenas 1 pedido registrado.</p>';
                                    }
                                    
                                    elementoCard.innerHTML = cardsHtml;
                                } else {
                                    elementoCard.innerHTML = `
                                        <div class="text-center py-8">
                                            <div class="bg-gray-100 rounded-full h-16 w-16 flex items-center justify-center mx-auto mb-3">
                                                <i class="fas fa-inbox text-gray-400 text-xl"></i>
                                            </div>
                                            <p class="text-gray-500 italic text-sm">Nenhum pedido encontrado para esta empresa.</p>
                                        </div>
                                    `;
                                }
                            })
                            .withFailureHandler(function(erro) {
                                console.error(`Erro ao buscar pedidos da empresa ${empresaAtual.nome}:`, erro);
                                var elementoCard = document.getElementById(`card-content-${empresaAtual.id}`);
                                elementoCard.innerHTML = `
                                    <div class="text-center py-8">
                                        <div class="bg-red-100 rounded-full h-16 w-16 flex items-center justify-center mx-auto mb-3">
                                            <i class="fas fa-exclamation-triangle text-red-400 text-xl"></i>
                                        </div>
                                        <p class="text-red-500 italic text-sm">Erro ao carregar pedidos desta empresa.</p>
                                    </div>
                                `;
                            })
                            .buscarPedidos('', empresaAtual.id);
                    });
                } else {
                    contentArea.innerHTML = '<div class="bg-white p-6 rounded-xl shadow-lg text-center"><p class="text-gray-500 italic">Nenhuma empresa encontrada para este usuário.</p></div>';
                }
            })
            .withFailureHandler(function(erro) {
                console.error("Erro ao buscar empresas:", erro);
                contentArea.innerHTML = '<div class="bg-white p-6 rounded-xl shadow-lg text-center"><p class="text-red-500 italic">Erro ao carregar lista de empresas.</p></div>';
            })
            .obterEmpresasDoUsuario(usuarioLogado);
        }
        
        // Fechamento da função setupMenuScreen
        }
        
        // ===============================================
        // FUNÇÕES ESPECÍFICAS PARA ADMIN NO DASHBOARD
        // ===============================================

        function abrirImpressaoPedidoAdmin(numeroPedido, empresaId, nomeEmpresa) {
              if (!numeroPedido || !empresaId) {
                  showGlobalModal('Erro', 'Dados incompletos para abrir o pedido.');
                  return;
              }
              
              console.log(`Admin abrindo pedido ${numeroPedido} da empresa ${nomeEmpresa} (ID: ${empresaId})`);
              showGlobalModal('Aguarde', `Preparando impressão do pedido ${numeroPedido} da empresa ${nomeEmpresa}...`, 'info');

              google.script.run
                  .withSuccessHandler(function(pedidoData) {
                      if (!pedidoData) {
                          showGlobalModal('Erro', 'Dados do pedido não encontrados para impressão.');
                          return;
                      }
                      
                      // Salvar temporariamente os dados da empresa para impressão
                      const empresaTemp = {
                          id: empresaId,
                          nome: nomeEmpresa
                      };
                      
                      abrirJanelaImpressaoAdmin(pedidoData, empresaTemp);
                  })
                  .withFailureHandler(err => {
                      showGlobalModal('Erro', `Não foi possível carregar os dados do pedido: ${err.message}`);
                      console.error("Erro ao buscar pedido admin:", err);
                  })
                  .getDadosPedidoParaImpressaoAdmin(numeroPedido, empresaId); // Nova função que aceita empresa
          }

          function imprimirPedidoAdmin(numeroPedido, empresaId, nomeEmpresa) {
              // Mesma lógica que abrirImpressaoPedidoAdmin
              abrirImpressaoPedidoAdmin(numeroPedido, empresaId, nomeEmpresa);
          }

          function abrirJanelaImpressaoAdmin(pedidoData, empresaTemp) {
              document.getElementById('global-modal').classList.add('hidden');

              const pageHtml = getPageTemplate('PedidoImpressao');
              const printWindow = window.open('', '_blank');
              
              if (!printWindow) {
                  showGlobalModal("Pop-up bloqueado!", "Por favor, permita pop-ups para este site para poder imprimir o pedido.");
                  return;
              }

              printWindow.document.write(pageHtml);
              printWindow.document.close();

              printWindow.onload = function() {
                  const doc = printWindow.document;
                  doc.title = `Pedido de Compra - ${pedidoData.numeroDoPedido}`;
                  
                  // Para admin, usar os dados da empresa passados como parâmetro
                  doc.getElementById('empresaNomePrint').textContent = empresaTemp.nome || '';
                  doc.getElementById('empresaEnderecoPrint').textContent = pedidoData.enderecoEmpresa || '';
                  doc.getElementById('empresaCidadeUfPrint').textContent = pedidoData.cidadeUfEmpresa || '';
                  doc.getElementById('empresaCnpjPrint').textContent = pedidoData.cnpjEmpresa || '';
                  doc.getElementById('empresaEmailPrint').textContent = `Email: ${pedidoData.emailEmpresa || ''}`;
                  doc.getElementById('empresaTelefonePrint').textContent = `Telefone: ${pedidoData.telefoneEmpresa || ''}`;

                  doc.getElementById('numeroPedidoPrint').textContent = pedidoData.numeroDoPedido || '';
                  
                  const dataPedidoObj = new Date(pedidoData.data + 'T12:00:00');
                  const formattedDate = dataPedidoObj.toLocaleDateString('pt-BR', { year: 'numeric', month: '2-digit', day: '2-digit' });
                  const now = new Date();
                  const formattedTime = now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
                  doc.getElementById('dataEmissaoPrint').textContent = `${formattedDate} ${formattedTime}`;

                  doc.getElementById('fornecedorPrint').textContent = pedidoData.fornecedor || '';
                  doc.getElementById('enderecoFornecedorPrint').textContent = pedidoData.enderecoFornecedor || '';
                  doc.getElementById('formaPagamentoPrint').textContent = pedidoData.formaPagamentoFornecedor || '';
                  doc.getElementById('cnpjFornecedorPrint').textContent = pedidoData.cnpjFornecedor || '';
                  doc.getElementById('condicaoPagamentoPrint').textContent = pedidoData.condicaoPagamentoFornecedor || '';
                  
                  doc.getElementById('placaVeiculoPrint').textContent = pedidoData.placaVeiculo || 'N/A';
                  doc.getElementById('nomeVeiculoPrint').textContent = pedidoData.nomeVeiculo || 'N/A';
                  
                  doc.getElementById('totalGeralPrint').textContent = `R$ ${parseFloat(pedidoData.totalGeral || 0).toFixed(2).replace('.', ',')}`;
                  doc.getElementById('observacoesPedidoPrint').textContent = pedidoData.observacoes || 'Sem observações.';

                  const printItemsTableBody = doc.getElementById('printItemsTableBody');
                  printItemsTableBody.innerHTML = '';
                  if (pedidoData.itens && Array.isArray(pedidoData.itens)) {
                      pedidoData.itens.forEach((item, index) => {
                          const row = printItemsTableBody.insertRow();
                          row.insertCell(0).textContent = index + 1;
                          row.insertCell(1).textContent = item.descricao;
                          row.insertCell(2).textContent = item.unidade || '';
                          row.insertCell(3).textContent = item.quantidade;
                          row.insertCell(4).textContent = `R$ ${parseFloat(item.precoUnitario || 0).toFixed(2).replace('.', ',')}`;
                          row.insertCell(5).textContent = `R$ ${parseFloat(item.totalItem || 0).toFixed(2).replace('.', ',')}`;
                      });
                  }
                  
                  const nomeUsuario = localStorage.getItem('nome') || 'Usuário Desconhecido';
                  const perfilUsuario = localStorage.getItem('perfil') || 'Sem Função';
                  doc.getElementById('usuarioLogadoPrint').textContent = nomeUsuario.toUpperCase();
                  doc.getElementById('funcaoUsuarioLogadoPrint').textContent = perfilUsuario.toUpperCase();
                };
          }

        // ===============================================
        // VISUALIZAR PEDIDO USANDO TELA EXISTENTE
        // ===============================================

        function visualizarPedidoExistente(numeroPedido, empresaId, nomeEmpresa) {
            console.log(`🔍 Visualizando pedido: ${numeroPedido}`);
            
            // Definir que veio do menu de usuário normal
            localStorage.setItem('modoVisualizacao', 'true');
            localStorage.setItem('origemVisualizacao', 'menu-usuario');
            localStorage.setItem('numeroPedidoVisualizacao', numeroPedido);
            
            // PRIMEIRO: Buscar dados do pedido
            console.log('🔍 Buscando dados do pedido antes de carregar a tela...');
            google.script.run
                .withSuccessHandler(async (response) => {
                    if (response.status === 'success' && response.data.length > 0) {
                        console.log("📄 Dados do pedido recebidos:", response.data);
                        
                        // Encontrar o pedido específico
                        let pedidoEncontrado = response.data.find(p => p.numeroDoPedido == numeroPedido);
                        
                        if (!pedidoEncontrado) {
                            // Tentar com string
                            pedidoEncontrado = response.data.find(p => String(p.numeroDoPedido) === String(numeroPedido));
                        }
                        
                        if (!pedidoEncontrado) {
                            // Tentar com padding
                            const numeroComZeros = numeroPedido.padStart(4, '0');
                            pedidoEncontrado = response.data.find(p => 
                                String(p.numeroDoPedido).padStart(4, '0') === numeroComZeros
                            );
                        }
                        
                        if (pedidoEncontrado) {
                            console.log("✅ Pedido encontrado:", pedidoEncontrado);
                            console.log("🔍 DEBUG - Campo estadoFornecedor:", pedidoEncontrado.estadoFornecedor);
                            console.log("🔍 DEBUG - Todos os campos do pedido:", Object.keys(pedidoEncontrado));
                            
                            // Salvar dados do pedido no localStorage ANTES de carregar a tela
                            localStorage.setItem('pedidoParaVisualizar', JSON.stringify(pedidoEncontrado));
                            console.log('✅ Dados salvos no localStorage');
                            
                            // AGORA carregar a tela de pedido
                            loadPageContent('Pedido', setupPedidoScreen);
                            
                        } else {
                            console.error("❌ Pedido não encontrado");
                            showGlobalModal("Erro", `Pedido ${numeroPedido} não encontrado.`);
                        }
                        
                    } else {
                        showGlobalModal("Erro", "Pedido não encontrado");
                    }
                })
                .withFailureHandler((error) => {
                    console.error("❌ Erro ao buscar pedido:", error);
                    showGlobalModal("Erro", "Erro ao carregar pedido: " + error.message);
                })
                .buscarPedidos(numeroPedido, empresaId);
        }

        // Função específica para visualização no dashboard admin
        function visualizarPedidoAdmin(numeroPedido, empresaId, nomeEmpresa) {
            console.log(`🔍 Admin visualizando pedido: ${numeroPedido} da empresa ${nomeEmpresa}`);
            
            // Definir que veio do dashboard admin
            localStorage.setItem('modoVisualizacao', 'true');
            localStorage.setItem('origemVisualizacao', 'dashboard-admin');
            localStorage.setItem('numeroPedidoVisualizacao', numeroPedido);
            localStorage.setItem('empresaOriginalAdmin', JSON.stringify({
                id: empresaId,
                nome: nomeEmpresa
            }));
            
            // PRIMEIRO: Buscar dados do pedido
            console.log('🔍 Buscando dados do pedido antes de carregar a tela...');
            google.script.run
                .withSuccessHandler(async (response) => {
                    if (response.status === 'success' && response.data.length > 0) {
                        console.log("📄 Dados do pedido recebidos (Admin):", response.data);
                        
                        // Encontrar o pedido específico
                        let pedidoEncontrado = response.data.find(p => p.numeroDoPedido == numeroPedido);
                        
                        if (!pedidoEncontrado) {
                            // Tentar com string
                            pedidoEncontrado = response.data.find(p => String(p.numeroDoPedido) === String(numeroPedido));
                        }
                        
                        if (!pedidoEncontrado) {
                            // Tentar com padding
                            const numeroComZeros = numeroPedido.padStart(4, '0');
                            pedidoEncontrado = response.data.find(p => 
                                String(p.numeroDoPedido).padStart(4, '0') === numeroComZeros
                            );
                        }
                        
                        if (pedidoEncontrado) {
                            console.log("✅ Pedido encontrado (Admin):", pedidoEncontrado);
                            console.log("🔍 DEBUG - Campo estadoFornecedor:", pedidoEncontrado.estadoFornecedor);
                            console.log("🔍 DEBUG - Todos os campos do pedido:", Object.keys(pedidoEncontrado));
                            
                            // Salvar dados do pedido no localStorage ANTES de carregar a tela
                            localStorage.setItem('pedidoParaVisualizar', JSON.stringify(pedidoEncontrado));
                            console.log('✅ Dados salvos no localStorage');
                            
                            // AGORA carregar a tela de pedido
                            loadPageContent('Pedido', setupPedidoScreen);
                            
                        } else {
                            console.error("❌ Pedido não encontrado (Admin)");
                            showGlobalModal("Erro", `Pedido ${numeroPedido} não encontrado.`);
                        }
                        
                    } else {
                        showGlobalModal("Erro", "Pedido não encontrado");
                    }
                })
                .withFailureHandler((error) => {
                    console.error("❌ Erro ao buscar pedido (Admin):", error);
                    showGlobalModal("Erro", "Erro ao carregar pedido: " + error.message);
                })
                .buscarPedidos(numeroPedido, empresaId);
        }

        function preencherTelaPedidoExistente() {
    try {
        console.log('=== INICIANDO PREENCHIMENTO DA TELA ===');
        
        // Recuperar dados salvos
        const pedidoDataString = localStorage.getItem('pedidoParaVisualizar');
        const empresaOriginal = localStorage.getItem('empresaOriginalAdmin');
        
        console.log('Dados do localStorage:');
        console.log('- pedidoParaVisualizar:', pedidoDataString);
        console.log('- empresaOriginalAdmin:', empresaOriginal);
        
        // Definir origem baseado na presença de empresaOriginalAdmin
        if (empresaOriginal) {
            console.log('🔍 Detectado: origem = dashboard-admin');
            localStorage.setItem('origemVisualizacao', 'dashboard-admin');
        } else {
            console.log('🔍 Detectado: origem = menu-usuario');
            localStorage.setItem('origemVisualizacao', 'menu-usuario');
        }
        
        if (!pedidoDataString) {
            console.error('❌ Dados do pedido não encontrados no localStorage');
            showGlobalModal('Erro', 'Dados do pedido não encontrados.');
            return;
        }
        
        let pedidoData;
        try {
            pedidoData = JSON.parse(pedidoDataString);
        } catch (parseError) {
            console.error('❌ Erro ao fazer parse dos dados:', parseError);
            showGlobalModal('Erro', 'Erro ao processar dados do pedido.');
            return;
        }
        
        console.log('Dados parseados:', pedidoData);
        
        if (!pedidoData || !pedidoData.numeroDoPedido) {
            console.error('❌ Dados do pedido inválidos:', pedidoData);
            showGlobalModal('Erro', 'Dados do pedido inválidos.');
            return;
        }
        
        console.log('✅ Dados válidos, iniciando preenchimento...');
        
        // Aguardar mais um pouco para garantir que a tela carregou
        setTimeout(() => {
            processarPreenchimentoTela(pedidoData, empresaOriginal);
        }, 1000);
        
    } catch (error) {
        console.error("❌ Erro no preenchimento da tela:", error);
        console.error("Stack trace:", error.stack);
        showGlobalModal('Erro', 'Erro ao preencher os dados na tela.');
    }
}

function processarPreenchimentoTela(pedidoData, empresaOriginal) {
    try {
        console.log('Verificando se elementos existem...');
        
        // Verificar se elementos principais existem
        const elementos = {
            numeroPedido: document.getElementById('numeroPedido'),
            dataPedido: document.getElementById('dataPedido'),
            fornecedorPedido: document.getElementById('fornecedorPedido'),
            form: document.querySelector('.max-w-7xl.mx-auto') || document.querySelector('main') || document.querySelector('body')
        };
        
        console.log('Elementos encontrados:', Object.keys(elementos).map(key => 
            `${key}: ${elementos[key] ? '✅' : '❌'}`
        ));
        
        // Verificar se pelo menos o formulário e o campo de número existem
        if (!elementos.form) {
            console.error('❌ Formulário não encontrado, tentando novamente...');
            setTimeout(() => {
                processarPreenchimentoTela(pedidoData, empresaOriginal);
            }, 1000);
            return;
        }
        
        if (!elementos.numeroPedido) {
            console.error('❌ Campo numeroPedido não encontrado, tentando novamente...');
            setTimeout(() => {
                processarPreenchimentoTela(pedidoData, empresaOriginal);
            }, 1000);
            return;
        }
        
        console.log('✅ Elementos principais encontrados, iniciando preenchimento...');
        
        // Preencher campos
        preencherCamposPedido(pedidoData);
        
        // Aguardar mais um pouco antes de desabilitar e adicionar botões
        setTimeout(() => {
            finalizarPreenchimento(pedidoData, empresaOriginal);
        }, 800);
        
    } catch (error) {
        console.error("❌ Erro no processamento:", error);
        showGlobalModal('Erro', 'Erro ao processar dados na tela.');
    }
}

function finalizarPreenchimento(pedidoData, empresaOriginal) {
    try {
        desabilitarEdicaoPedido();
        
        // Verificar origem antes de adicionar botões
        const origemVisualizacao = localStorage.getItem('origemVisualizacao');
        console.log('🔧 finalizarPreenchimento - origemVisualizacao:', origemVisualizacao);
        console.log('🔧 finalizarPreenchimento - empresaOriginal:', empresaOriginal);
        
        if (origemVisualizacao === 'dashboard-admin' && empresaOriginal) {
            console.log('🔧 Adicionando botões do dashboard admin');
            adicionarBotaoVoltarDashboard(empresaOriginal);
            // Também aplicar modo visualização para desabilitar campos
            setupModoVisualizacao();
        } else if (origemVisualizacao === 'menu-usuario') {
            console.log('🔧 Adicionando botões do menu usuário via finalizarPreenchimento');
            setupModoVisualizacao();
        } else {
            console.log('🔧 Origem não identificada ou sem botões a adicionar');
        }
        
        // Limpar dados temporários
        localStorage.removeItem('pedidoParaVisualizar');
        
        showToast(`Pedido ${pedidoData.numeroDoPedido} carregado com sucesso!`, 'success', 'Sucesso');
        
        console.log('✅ Preenchimento concluído com sucesso');
        
        // Garantir que os campos sejam desabilitados após o preenchimento completo
        setTimeout(() => {
            console.log('🔒 Garantindo desabilitação final dos campos...');
            
            // Usar seletor mais direto
            const inputs = document.querySelectorAll('input');
            const selects = document.querySelectorAll('select');
            const textareas = document.querySelectorAll('textarea');
            const buttons = document.querySelectorAll('button');
            
            console.log(`🔒 Desabilitação final: ${inputs.length} inputs, ${selects.length} selects, ${textareas.length} textareas, ${buttons.length} buttons`);
            
            [...inputs, ...selects, ...textareas].forEach(campo => {
                if (!campo.readOnly && !campo.disabled) {
                    campo.disabled = true;
                    campo.readOnly = true;
                    campo.classList.add('bg-gray-100', 'text-gray-600', 'cursor-not-allowed');
                    console.log(`🔒 Campo desabilitado (final): ${campo.id || campo.name || campo.tagName}`);
                }
            });
            
            // Desabilitar botões (exceto navegação)
            buttons.forEach(button => {
                const id = button.id || '';
                const isNavButton = id.includes('voltar') || id.includes('imprimir') || id.includes('Menu') || id.includes('Dashboard');
                
                if (!isNavButton) {
                    button.disabled = true;
                    button.classList.add('opacity-50', 'cursor-not-allowed');
                    console.log(`🔒 Botão desabilitado (final): ${button.id || button.textContent?.trim().substring(0, 20)}`);
                }
            });
        }, 500);
        
    } catch (error) {
        console.error("❌ Erro na finalização:", error);
        showGlobalModal('Erro', 'Erro ao finalizar o preenchimento.');
    }
}

function preencherCamposPedido(dadosPedido) {
    console.log("=== PREENCHENDO CAMPOS ===");
    console.log("Dados recebidos:", dadosPedido);
       
    // 1. NÚMERO DO PEDIDO
    const numeroPedido = document.getElementById('numeroPedido');
    if (numeroPedido) {
        const numeroParaPreencher = dadosPedido.numeroDoPedido || dadosPedido.numero || '';
        numeroPedido.value = numeroParaPreencher;
        console.log(`✅ Número do pedido preenchido: "${numeroParaPreencher}"`);
        
        // Forçar atualização visual
        numeroPedido.dispatchEvent(new Event('input'));
        numeroPedido.dispatchEvent(new Event('change'));
    } else {
        console.error('❌ Campo numeroPedido não encontrado no DOM');
        // Tentar novamente após delay
        setTimeout(() => {
            const numeroPedidoRetry = document.getElementById('numeroPedido');
            if (numeroPedidoRetry) {
                const numeroParaPreencher = dadosPedido.numeroDoPedido || dadosPedido.numero || '';
                numeroPedidoRetry.value = numeroParaPreencher;
                console.log(`✅ Número do pedido preenchido (retry): "${numeroParaPreencher}"`);
            }
        }, 500);
    }
    
    // 2. DATA DO PEDIDO
    const dataPedido = document.getElementById('dataPedido');
    if (dataPedido) {
        dataPedido.value = dadosPedido.data || '';
        console.log(`✅ Data: ${dadosPedido.data}`);
    } else {
        console.warn('❌ Campo dataPedido não encontrado');
    }
    
    // 3. PLACA DO VEÍCULO
    const placaVeiculo = document.getElementById('placaVeiculo');
    if (placaVeiculo) {
        placaVeiculo.value = dadosPedido.placaVeiculo || '';
        console.log(`✅ Placa: ${dadosPedido.placaVeiculo}`);
    } else {
        console.warn('❌ Campo placaVeiculo não encontrado');
    }
    
    // 4. OBSERVAÇÕES
    const observacoesPedido = document.getElementById('observacoesPedido');
    if (observacoesPedido) {
        observacoesPedido.value = dadosPedido.observacoes || '';
        console.log(`✅ Observações: ${dadosPedido.observacoes}`);
    } else {
        console.warn('❌ Campo observacoesPedido não encontrado');
    }
    
    // 5. TOTAL GERAL
    const totalGeral = document.getElementById('totalGeral');
    if (totalGeral) {
        const valorFormatado = `R$ ${parseFloat(dadosPedido.totalGeral || 0).toFixed(2).replace('.', ',')}`;
        totalGeral.value = valorFormatado;
        console.log(`✅ Total: ${valorFormatado}`);
    } else {
        console.warn('❌ Campo totalGeral não encontrado');
    }
    
    // 6. FORNECEDOR E VEÍCULO - aguardar carregamento
    setTimeout(() => {
        preencherSelectFornecedor(dadosPedido.fornecedor);
        preencherSelectVeiculo(dadosPedido.nomeVeiculo);
    }, 1500);
    
    // 7. ITENS
    if (dadosPedido.itens && Array.isArray(dadosPedido.itens)) {
        console.log(`Preenchendo ${dadosPedido.itens.length} itens...`);
        preencherItensContainer(dadosPedido.itens);
    } else {
        console.warn('❌ Nenhum item encontrado ou formato inválido');
    }
}

function preencherSelectFornecedor(nomeFornecedor) {
    if (!nomeFornecedor) {
        console.warn('❌ Nome do fornecedor não fornecido');
        return;
    }
    
    const fornecedorSelect = document.getElementById('fornecedorPedido');
    if (fornecedorSelect && fornecedorSelect.options.length > 0) {
        console.log(`Procurando fornecedor: "${nomeFornecedor}"`);
        console.log(`Opções disponíveis:`, Array.from(fornecedorSelect.options).map(opt => opt.text));
        
        const opcoes = fornecedorSelect.options;
        for (let i = 0; i < opcoes.length; i++) {
            if (opcoes[i].text.includes(nomeFornecedor) || 
                opcoes[i].value.includes(nomeFornecedor)) {
                fornecedorSelect.selectedIndex = i;
                console.log(`✅ Fornecedor selecionado: ${opcoes[i].text}`);
                
                // Disparar evento de mudança para atualizar o estado
                const evento = new Event('change', { bubbles: true });
                fornecedorSelect.dispatchEvent(evento);
                
                return;
            }
        }
        console.warn(`❌ Fornecedor "${nomeFornecedor}" não encontrado nas opções`);
    } else {
        console.warn('❌ Select de fornecedor não carregado ainda');
        // Tentar novamente após mais tempo
        setTimeout(() => {
            preencherSelectFornecedor(nomeFornecedor);
        }, 1000);
    }
}

function preencherSelectVeiculo(nomeVeiculo) {
    if (!nomeVeiculo || nomeVeiculo.trim() === '') {
        console.warn('❌ Nome do veículo não fornecido ou vazio');
        return;
    }
    
    const veiculoSelect = document.getElementById('nomeVeiculo');
    if (veiculoSelect && veiculoSelect.options.length > 1) { // >1 porque sempre tem a opção vazia
        console.log(`Procurando veículo: "${nomeVeiculo}"`);
        console.log(`Opções disponíveis:`, Array.from(veiculoSelect.options).map(opt => opt.value).filter(v => v.trim() !== ''));
        
        const nomeVeiculoLimpo = nomeVeiculo.trim().toUpperCase();
        const opcoes = veiculoSelect.options;
        
        // Primeira tentativa: busca exata
        for (let i = 0; i < opcoes.length; i++) {
            const opcaoTexto = opcoes[i].textContent.trim().toUpperCase();
            const opcaoValue = opcoes[i].value.trim().toUpperCase();
            
            if (opcaoTexto === nomeVeiculoLimpo || opcaoValue === nomeVeiculoLimpo) {
                veiculoSelect.selectedIndex = i;
                console.log(`✅ Veículo selecionado (exato): ${opcoes[i].textContent}`);
                return;
            }
        }
        
        // Segunda tentativa: busca por conteúdo
        for (let i = 0; i < opcoes.length; i++) {
            const opcaoTexto = opcoes[i].textContent.trim().toUpperCase();
            const opcaoValue = opcoes[i].value.trim().toUpperCase();
            
            if (opcaoTexto.includes(nomeVeiculoLimpo) || 
                opcaoValue.includes(nomeVeiculoLimpo) ||
                nomeVeiculoLimpo.includes(opcaoTexto) ||
                nomeVeiculoLimpo.includes(opcaoValue)) {
                veiculoSelect.selectedIndex = i;
                console.log(`✅ Veículo selecionado (parcial): ${opcoes[i].textContent}`);
                return;
            }
        }
        
        console.warn(`❌ Veículo "${nomeVeiculo}" não encontrado nas opções disponíveis`);
    } else {
        console.warn('❌ Select de veículo não carregado ainda ou sem opções válidas');
        // Tentar novamente após mais tempo
        setTimeout(() => {
            preencherSelectVeiculo(nomeVeiculo);
        }, 1000);
    }
}

// Função para obter data/hora local no formato correto
function obterDataHoraLocal() {
    const agora = new Date();
    const ano = agora.getFullYear();
    const mes = String(agora.getMonth() + 1).padStart(2, '0');
    const dia = String(agora.getDate()).padStart(2, '0');
    const hora = String(agora.getHours()).padStart(2, '0');
    const minuto = String(agora.getMinutes()).padStart(2, '0');
    const segundo = String(agora.getSeconds()).padStart(2, '0');
    
    return `${ano}-${mes}-${dia} ${hora}:${minuto}:${segundo}`;
}

function preencherItensContainer(itens) {
    console.log("Preenchendo tabela de itens com", itens.length, "itens");
    
    const tabelaBody = document.getElementById('itensTableBody');
    if (!tabelaBody) {
        console.error("❌ Tabela de itens não encontrada");
        return;
    }
    
    // Limpar tabela (remover linha "nenhum item")
    tabelaBody.innerHTML = '';
    
    // Adicionar cada item como linha da tabela
    itens.forEach((item, index) => {
        const row = document.createElement('tr');
        row.className = 'hover:bg-gray-50';
        
        const precoFormatado = `R$ ${parseFloat(item.precoUnitario || 0).toFixed(2).replace('.', ',')}`;
        const totalFormatado = `R$ ${parseFloat(item.totalItem || 0).toFixed(2).replace('.', ',')}`;
        
        row.innerHTML = `
            <td class="px-6 py-4 text-sm text-gray-900">${item.descricao || ''}</td>
            <td class="px-6 py-4 text-sm text-gray-900 text-center">${item.quantidade || ''}</td>
            <td class="px-6 py-4 text-sm text-gray-900 text-center">${item.unidade || ''}</td>
            <td class="px-6 py-4 text-sm text-gray-900 text-right">${precoFormatado}</td>
            <td class="px-6 py-4 text-sm font-bold text-blue-600 text-right">${totalFormatado}</td>
            <td class="px-6 py-4 text-center">
                <button onclick="removerItem(${index})" class="text-red-600 hover:text-red-800 text-sm" title="Remover item">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        
        tabelaBody.appendChild(row);
    });
    
    console.log(`✅ Tabela preenchida com ${itens.length} itens`);
    
    // Calcular e atualizar total geral após carregar itens
    setTimeout(() => {
        const totalCalculado = itens.reduce((total, item) => {
            return total + (parseFloat(item.totalItem || 0));
        }, 0);
        
        const totalGeralElement = document.getElementById('totalGeral');
        const summaryTotalElement = document.getElementById('summary-total');
        
        if (totalGeralElement) {
            const valorFormatado = `R$ ${totalCalculado.toFixed(2).replace('.', ',')}`;
            totalGeralElement.textContent = valorFormatado;
            console.log(`✅ Total geral recalculado: ${valorFormatado}`);
        }
        
        if (summaryTotalElement) {
            const valorFormatado = `R$ ${totalCalculado.toFixed(2).replace('.', ',')}`;
            summaryTotalElement.textContent = valorFormatado;
            console.log(`✅ Summary total atualizado: ${valorFormatado}`);
        }
    }, 100);
}

        // Função para remover item da lista
        function removerItem(indexOuId) {
            if (confirm('Tem certeza que deseja remover este item?')) {
                console.log('🗑️ Removendo item:', indexOuId);
                
                // Se é um número simples (índice), remover do array window.itensAdicionados
                if (typeof indexOuId === 'number' && window.itensAdicionados && window.itensAdicionados[indexOuId]) {
                    console.log('🗑️ Removendo item do array (modo edição):', window.itensAdicionados[indexOuId]);
                    
                    // Remover do array
                    window.itensAdicionados.splice(indexOuId, 1);
                    
                    // Atualizar a tabela
                    preencherItensContainer(window.itensAdicionados);
                    
                    // Recalcular e atualizar total
                    const novoTotal = window.itensAdicionados.reduce((total, item) => {
                        return total + (parseFloat(item.subtotal || item.totalItem || 0));
                    }, 0);
                    
                    const totalGeralElement = document.getElementById('totalGeral');
                    const summaryTotalElement = document.getElementById('summary-total');
                    
                    if (totalGeralElement) {
                        totalGeralElement.textContent = `R$ ${novoTotal.toFixed(2).replace('.', ',')}`;
                    }
                    
                    if (summaryTotalElement) {
                        summaryTotalElement.textContent = `R$ ${novoTotal.toFixed(2).replace('.', ',')}`;
                    }
                    
                    showToast('Item removido com sucesso!', 'success');
                    console.log('✅ Item removido. Total atualizado:', novoTotal);
                } 
                // Se é um ID específico (novo item), remover da tabela diretamente
                else {
                    const row = document.querySelector(`tr[data-item-id="${indexOuId}"]`);
                    if (row) {
                        console.log('🗑️ Removendo linha da tabela (novo item):', row);
                        
                        // Remover do array global se existir
                        if (window.itensAdicionados) {
                            const itemIndex = window.itensAdicionados.findIndex(item => item.id === indexOuId);
                            if (itemIndex !== -1) {
                                window.itensAdicionados.splice(itemIndex, 1);
                            }
                        }
                        
                        row.remove();
                        calcularTotalGeral();
                        showToast('Item removido com sucesso!', 'success');
                    } else {
                        console.warn('⚠️ Linha com ID não encontrada:', indexOuId);
                    }
                }
            }
        }

        function desabilitarEdicaoPedido() {
          // Desabilitar todos os campos do formulário
          const form = document.querySelector('.max-w-7xl.mx-auto') || document.querySelector('main');
          if (form) {
              const inputs = form.querySelectorAll('input:not([readonly]), select, textarea');
              inputs.forEach(input => {
                  input.disabled = true;
                  input.classList.add('bg-gray-100', 'text-gray-600');
              });
              
              // Desabilitar botões específicos mas manter alguns
              const botoes = form.querySelectorAll('button');
              botoes.forEach(botao => {
                  const id = botao.id || '';
                  const texto = botao.textContent.toLowerCase();
                  
                  // Manter apenas botões permitidos
                  if (!id.includes('voltar') && 
                      !id.includes('imprimir') && 
                      !texto.includes('voltar') && 
                      !texto.includes('imprimir')) {
                      botao.disabled = true;
                      botao.classList.add('opacity-50', 'cursor-not-allowed');
                  }
              });
              
              console.log('✅ Modo visualização ativado - campos desabilitados');
          }
      }


        function adicionarBotaoVoltarDashboard(empresaOriginal) {
          console.log('🔧 adicionarBotaoVoltarDashboard chamado - empresaOriginal:', empresaOriginal);
          
          const form = document.querySelector('.max-w-7xl.mx-auto') || document.querySelector('main');
          if (!form) {
              console.warn('Container principal não encontrado para adicionar botões');
              return;
          }
          
          // Verificar se já existe
              if (document.getElementById('botoes-admin-visualizacao')) {
                  console.log('⚠️ Botões de admin já existem - não adicionando novamente');
                  return;
              }
              
              console.log('🔧 Adicionando botões de admin para visualização');
              
              // Encontrar a div dos botões existentes
              const botoesExistentes = form.querySelector('.flex.justify-end.space-x-4');
              if (botoesExistentes) {
                  console.log('🔧 Botões existentes encontrados - escondendo');
                  // Esconder botões originais (salvar/cancelar)
                  botoesExistentes.style.display = 'none';
                  
                  // Adicionar novos botões
                  const novosBotoes = document.createElement('div');
                  novosBotoes.id = 'botoes-admin-visualizacao';
                  novosBotoes.className = 'flex justify-end space-x-4 pt-4 border-t border-gray-200';
                  novosBotoes.innerHTML = `
                      <div class="flex items-center text-sm text-gray-500 mr-4">
                          <i class="fas fa-eye mr-2"></i>
                          Visualizando em modo somente leitura (Admin)
                      </div>
                      <button id="btn-voltar-dashboard-admin" 
                              class="px-6 py-2 bg-gray-500 hover:bg-gray-600 text-white font-semibold rounded-lg shadow-md transition-colors duration-200">
                          <i class="fas fa-arrow-left mr-2"></i>Voltar ao Dashboard
                      </button>
                      <button id="btn-imprimir-pedido-atual" 
                              class="px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white font-semibold rounded-lg shadow-md transition-colors duration-200">
                          <i class="fas fa-print mr-2"></i>Imprimir Pedido
                      </button>
                  `;
                  
                  // Inserir após os botões originais
                  botoesExistentes.parentNode.insertBefore(novosBotoes, botoesExistentes.nextSibling);
                  
                  // Adicionar event listeners aos novos botões
                  const btnVoltar = document.getElementById('btn-voltar-dashboard-admin');
                  const btnImprimir = document.getElementById('btn-imprimir-pedido-atual');
                  
                  if (btnVoltar) {
                      btnVoltar.addEventListener('click', function(e) {
                          e.preventDefault();
                          e.stopPropagation();
                          voltarDashboardAdmin(empresaOriginal || '');
                      });
                  }
                  
                  if (btnImprimir) {
                      btnImprimir.addEventListener('click', function(e) {
                          e.preventDefault();
                          e.stopPropagation();
                          imprimirPedidoAtual();
                      });
                  }
                  
                  console.log('✅ Botões de admin adicionados com event listeners');
              }
          }

        // ===============================================
        // FUNÇÕES AUXILIARES PARA VISUALIZAÇÃO DE PEDIDOS
        // ===============================================
        
        function limparEstadoVisualizacao() {
            try {
                // Limpar dados temporários
                localStorage.removeItem('empresaOriginalAdmin');
                localStorage.removeItem('pedidoParaVisualizar');
                localStorage.removeItem('modoVisualizacao');
                
                // Remover botões de admin se existirem
                const botoesAdmin = document.getElementById('botoes-admin-visualizacao');
                if (botoesAdmin) {
                    botoesAdmin.remove();
                }
                
                console.log('Estado de visualização limpo');
            } catch (error) {
                console.error('Erro ao limpar estado de visualização:', error);
            }
        }
        
        function verificarElementosNecessarios() {
            const elementos = {
                numeroPedido: document.getElementById('numeroPedido'),
                globalModal: document.getElementById('global-modal'),
                mainContent: document.getElementById('main-content-area')
            };
            
            const elementosValidos = Object.keys(elementos).filter(key => elementos[key] !== null);
            console.log('Elementos disponíveis:', elementosValidos);
            
            return elementos;
        }

        function voltarDashboardAdmin(empresaOriginal) {
            try {
                console.log('Voltando ao Menu...');
                
                // Limpar estado de visualização primeiro
                limparEstadoVisualizacao();
                
                // Verificar elementos necessários
                const elementos = verificarElementosNecessarios();
                
                // Restaurar empresa original se existir
                if (empresaOriginal && empresaOriginal !== 'null' && empresaOriginal !== 'undefined') {
                    try {
                        // Verificar se é um JSON válido
                        if (typeof empresaOriginal === 'string' && empresaOriginal.startsWith('{')) {
                            const empresaJson = JSON.parse(empresaOriginal);
                            if (empresaJson && empresaJson.id) {
                                localStorage.setItem('empresaSelecionada', empresaOriginal);
                                console.log('Empresa original restaurada');
                            }
                        }
                    } catch (e) {
                        console.warn('Erro ao restaurar empresa original:', e);
                    }
                }
                
                // Verificar se a função loadPageContent existe
                if (typeof loadPageContent !== 'function') {
                    console.error('Função loadPageContent não encontrada');
                    location.reload(); // Fallback extremo
                    return;
                }
                
                // Verificar se setupMenuScreen existe
                if (typeof setupMenuScreen !== 'function') {
                    console.error('Função setupMenuScreen não encontrada');
                    location.reload(); // Fallback extremo
                    return;
                }
                
                // Voltar ao Menu com timeout para evitar conflitos
                setTimeout(() => {
                    try {
                        loadPageContent('Menu', setupMenuScreen);
                    } catch (loadError) {
                        console.error('Erro ao carregar menu:', loadError);
                        // Fallback - recarregar página
                        location.reload();
                    }
                }, 200);
                
            } catch (error) {
                console.error('Erro em voltarDashboardAdmin:', error);
                // Fallback final - limpar tudo e recarregar
                limparEstadoVisualizacao();
                setTimeout(() => {
                    location.reload();
                }, 300);
            }
        }

        function imprimirPedidoAtual() {
            try {
                console.log('=== INICIANDO IMPRESSÃO DO PEDIDO ATUAL ===');
                
                // Verificar elementos necessários
                const elementos = verificarElementosNecessarios();
                
                const numeroPedido = elementos.numeroPedido?.value?.trim();
                
                if (!numeroPedido) {
                    console.error('Número do pedido não encontrado');
                    showGlobalModal('Erro', 'Número do pedido não encontrado para impressão.');
                    return;
                }
                
                console.log('Número do pedido encontrado:', numeroPedido);
                
                // Verificar se a função de impressão existe
                if (typeof abrirImpressaoPedido !== 'function') {
                    console.error('Função abrirImpressaoPedido não encontrada');
                    showGlobalModal('Erro', 'Função de impressão não disponível.');
                    return;
                }
                
                // Verificar conexão com Google Apps Script
                if (typeof google === 'undefined' || !google.script) {
                    console.error('Google Apps Script não disponível');
                    showGlobalModal('Erro', 'Conexão com servidor não disponível.');
                    return;
                }
                
                // Chamar função de impressão com timeout para evitar conflitos
                setTimeout(() => {
                    try {
                        abrirImpressaoPedido(numeroPedido);
                    } catch (printError) {
                        console.error('Erro na função de impressão:', printError);
                        showGlobalModal('Erro', `Erro ao executar impressão: ${printError.message}`);
                    }
                }, 150);
                
            } catch (error) {
                console.error('Erro em imprimirPedidoAtual:', error);
                showGlobalModal('Erro', `Erro ao imprimir pedido: ${error.message}`);
            }
        }

        // ===============================================
        // FUNÇÃO PARA GERENCIAR USUÁRIOS
        // ===============================================
        
        function setupGerenciarUsuariosScreen() {
            // Dados que serão carregados do backend
            let allUsers = [];
            let empresasDisponiveis = [];
            let auditDataCache = {};

            // ===============================================
            // FUNÇÕES DE CARREGAMENTO DE DADOS
            // ===============================================

            function loadUsersFromBackend() {
                showToast('Carregando usuários...', 'info');
                
                google.script.run
                    .withSuccessHandler(users => {
                        allUsers = users || [];
                        renderUsersTable();
                        showToast('Usuários carregados com sucesso!', 'success');
                    })
                    .withFailureHandler(error => {
                        console.error('Erro ao carregar usuários:', error);
                        showToast('Erro ao carregar usuários: ' + error.message, 'error');
                    })
                    .listarUsuariosCompleto();
            }

            function loadEmpresasFromBackend() {
                google.script.run
                    .withSuccessHandler(empresas => {
                        empresasDisponiveis = empresas.map(emp => emp.nome) || [];
                    })
                    .withFailureHandler(error => {
                        console.error('Erro ao carregar empresas:', error);
                        empresasDisponiveis = ['Matriz', 'Filial SP', 'Filial RJ']; // Fallback
                    })
                    .listarEmpresas();
            }

            function loadAuditData(userId, callback) {
                if (auditDataCache[userId]) {
                    callback(auditDataCache[userId]);
                    return;
                }

                google.script.run
                    .withSuccessHandler(auditData => {
                        auditDataCache[userId] = auditData;
                        callback(auditData);
                    })
                    .withFailureHandler(error => {
                        console.error('Erro ao carregar dados de auditoria:', error);
                        callback({ lastLogin: null, lastOrder: null, lastPrint: null });
                    })
                    .obterDadosAuditoria(userId);
            }
            // ===============================================
            // ELEMENTOS DO DOM
            // ===============================================
            const tableBody = document.getElementById('usersTableBody');
            const messageElement = document.getElementById('userManagementMessage');
            // Modais
            const modals = {
                empresas: document.getElementById('modal-empresas-container'),
                audit: document.getElementById('modal-audit-container')
            };

            // ===============================================
            // FUNÇÕES DE RENDERIZAÇÃO E UI
            // ===============================================

            function renderUsersTable() {
                if (!tableBody) return;
                tableBody.innerHTML = '';

                allUsers.forEach(user => {
                    const empresasArray = user.empresas ? user.empresas.split(',').filter(e => e.trim()) : [];
                    const empresasHtml = empresasArray.map(empresa => {
                        const isDefault = empresa.trim() === user.empresaPadrao;
                        return `<span class="inline-flex items-center ${isDefault ? 'tag-padrao' : 'bg-gray-100 text-gray-600'} text-xs font-medium mr-2 px-2.5 py-1 rounded-full border border-transparent">${isDefault ? '<i class="fas fa-star fa-xs mr-1.5 text-yellow-500"></i>' : ''}${empresa.trim()}</span>`;
                    }).join('') || `<span class="text-xs text-gray-400 italic">Nenhuma</span>`;

                    const statusClass = user.status === 'Ativo' ? 'status-ativo' : 'status-inativo';
                    const toggleIcon = user.status === 'Ativo' ? 'fa-toggle-on text-green-500' : 'fa-toggle-off text-gray-400';

                    const row = document.createElement('tr');
                    row.className = 'hover:bg-gray-50';
                    row.innerHTML = `
                        <td class="px-6 py-4 whitespace-nowrap"><div class="text-sm font-medium text-gray-900">${user.nome}</div><div class="text-sm text-gray-500">${user.usuario}</div></td>
                        <td class="px-6 py-4 whitespace-nowrap"><span class="px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${statusClass}">${user.status}</span></td>
                        <td class="px-6 py-4 text-sm text-gray-500"><div class="flex flex-wrap gap-2">${empresasHtml}</div></td>
                        <td class="px-6 py-4 whitespace-nowrap text-center text-xl font-medium space-x-5">
                            <button data-userid="${user.id}" class="action-toggle-status text-gray-400 hover:text-gray-600" title="${user.status === 'Ativo' ? 'Desativar' : 'Ativar'}"><i class="fas ${toggleIcon}"></i></button>
                            <button data-userid="${user.id}" class="action-manage-empresas text-gray-400 hover:text-blue-600" title="Gerenciar Empresas"><i class="fas fa-building"></i></button>
                            <button data-userid="${user.id}" class="action-audit text-gray-400 hover:text-blue-600" title="Auditoria"><i class="fas fa-shield-halved"></i></button>
                        </td>
                    `;
                    tableBody.appendChild(row);
                });
                addTableEventListeners();
            }

            function showMessage(text, type = 'success') {
                messageElement.textContent = text;
                messageElement.className = `text-center text-sm font-medium min-h-[20px] mb-4 transition-opacity duration-300 ${type === 'success' ? 'text-green-600' : 'text-red-600'}`;
                setTimeout(() => { messageElement.textContent = ''; }, 3000);
            }

            // ===============================================
            // LÓGICA DOS MODAIS
            // ===============================================

            function openModal(modalName, userId) {
                const user = allUsers.find(u => u.id == userId);
                if (!user) return;
                
                const modalContainer = modals[modalName];
                if (!modalContainer) return;

                if (modalName === 'empresas') {
                    modalContainer.querySelector('#modal-empresas-title').textContent = `Empresas de: ${user.nome}`;
                    const listBody = modalContainer.querySelector('#empresas-list-body');
                    
                    // Carregamos empresas disponíveis e as permissões do usuário
                    google.script.run
                        .withSuccessHandler(empresas => {
                            const userEmpresasIds = user.empresasIds || [];
                            listBody.innerHTML = empresas.map(empresa => `
                                <tr class="hover:bg-gray-50 border-b border-gray-100">
                                    <td class="px-4 py-3 text-center"><input type="checkbox" value="${empresa.codigo}" name="empresa_access" class="h-5 w-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500" ${userEmpresasIds.includes(String(empresa.codigo)) ? 'checked' : ''}></td>
                                    <td class="px-4 py-3 text-sm text-gray-800">${empresa.nome}</td>
                                    <td class="px-4 py-3 text-center"><input type="radio" value="${empresa.codigo}" name="empresa_default" class="h-5 w-5 text-blue-600 focus:ring-blue-500" ${user.empresaPadraoId === String(empresa.codigo) ? 'checked' : ''}></td>
                                </tr>`).join('');
                        })
                        .withFailureHandler(error => {
                            console.error('Erro ao carregar empresas:', error);
                            listBody.innerHTML = '<tr><td colspan="3" class="text-center text-red-500">Erro ao carregar empresas</td></tr>';
                        })
                        .listarEmpresas();
                    
                    modalContainer.querySelector('#modal-empresas-save-btn').dataset.userid = userId;
                }

                if (modalName === 'audit') {
                    modalContainer.querySelector('#modal-audit-title').textContent = `Auditoria de: ${user.nome}`;
                    const auditBody = modalContainer.querySelector('#modal-audit-body');
                    
                    // Mostra loading
                    auditBody.innerHTML = '<div class="text-center text-gray-500">Carregando dados de auditoria...</div>';
                    
                    loadAuditData(userId, (data) => {
                        const format = (d) => d ? new Date(d).toLocaleString('pt-BR') : 'N/A';
                        
                        const createAuditItem = (icon, title, value) => `<div class="flex items-start"><div class="flex-shrink-0 w-8 text-center"><i class="fas ${icon} text-gray-400"></i></div><div class="ml-3"><p class="text-sm font-medium text-gray-800">${title}</p><p class="text-sm text-gray-500">${value}</p></div></div>`;
                        
                        auditBody.innerHTML = 
                            createAuditItem('fa-clock', 'Último Acesso', `${format(data.lastLogin?.date)} (IP: ${data.lastLogin?.ip || 'N/A'})`) +
                            createAuditItem('fa-file-invoice', 'Último Pedido Criado', data.lastOrder ? `${data.lastOrder.id} em ${format(data.lastOrder.date)}` : 'Nenhum registro') +
                            createAuditItem('fa-print', 'Última Impressão', data.lastPrint ? `Pedido ${data.lastPrint.id} em ${format(data.lastPrint.date)}` : 'Nenhum registro');
                    });
                }
                
                modalContainer.classList.remove('modal-hidden');
            }

            function closeModal(modalContainer) {
                modalContainer.classList.add('modal-hidden');
            }

            // ===============================================
            // AÇÕES DA TABELA
            // ===============================================
            
            function toggleUserStatus(userId) {
                const user = allUsers.find(u => u.id == userId);
                if (!user) return;
                
                showToast('Alterando status do usuário...', 'info');
                
                google.script.run
                    .withSuccessHandler(response => {
                        if (response.status === 'ok') {
                            showToast(response.message, 'success');
                            // Recarrega os dados para manter sincronia
                            loadUsersFromBackend();
                        } else {
                            showToast('Erro: ' + response.message, 'error');
                        }
                    })
                    .withFailureHandler(error => {
                        console.error('Erro ao alterar status:', error);
                        showToast('Erro ao alterar status: ' + error.message, 'error');
                    })
                    .alternarStatusUsuario(userId);
            }
            
            function saveEmpresasChanges(btn) {
                const userId = btn.dataset.userid;
                const user = allUsers.find(u => u.id == userId);
                if (!user) return;
                
                const modalContainer = modals.empresas;
                const empresasSelecionadas = Array.from(modalContainer.querySelectorAll('input[name="empresa_access"]:checked')).map(cb => cb.value);
                const defaultRadio = modalContainer.querySelector('input[name="empresa_default"]:checked');
                const empresaPadrao = defaultRadio ? defaultRadio.value : '';
                
                showToast('Salvando permissões de empresas...', 'info');
                
                google.script.run
                    .withSuccessHandler(response => {
                        if (response.status === 'ok' || response.status === 'success') {
                            showToast(`Empresas de ${user.nome} atualizadas com sucesso!`, 'success', 'Permissões Atualizadas');
                            closeModal(modalContainer);
                            // Recarrega os dados para mostrar as mudanças
                            loadUsersFromBackend();
                        } else {
                            showToast('Erro: ' + response.message, 'error', 'Falha na Atualização');
                        }
                    })
                    .withFailureHandler(error => {
                        console.error('Erro ao salvar permissões:', error);
                        showToast('Erro ao salvar permissões: ' + error.message, 'error');
                    })
                    .salvarPermissoesEmpresaUsuario(userId, empresasSelecionadas, empresaPadrao);
            }

            // ===============================================
            // EVENT LISTENERS
            // ===============================================

            function addTableEventListeners() {
                document.querySelectorAll('.action-toggle-status, .action-manage-empresas, .action-audit').forEach(button => {
                    button.addEventListener('click', (e) => {
                        const target = e.currentTarget;
                        const userId = target.dataset.userid;
                        if (target.classList.contains('action-toggle-status')) toggleUserStatus(userId);
                        if (target.classList.contains('action-manage-empresas')) openModal('empresas', userId);
                        if (target.classList.contains('action-audit')) openModal('audit', userId);
                    });
                });
            }
            
            // Listeners dos Modais
            document.querySelectorAll('.modal-container').forEach(modalContainer => {
                const overlay = modalContainer.querySelector('.modal-overlay');
                const closeBtn = modalContainer.querySelector('.modal-close-btn');
                const cancelBtn = modalContainer.querySelector('.modal-cancel-btn');
                
                if (overlay) overlay.addEventListener('click', () => closeModal(modalContainer));
                if (closeBtn) closeBtn.addEventListener('click', () => closeModal(modalContainer));
                if (cancelBtn) cancelBtn.addEventListener('click', () => closeModal(modalContainer));
            });
            
            if (modals.empresas) {
                const saveBtn = modals.empresas.querySelector('#modal-empresas-save-btn');
                if (saveBtn) saveBtn.addEventListener('click', (e) => saveEmpresasChanges(e.currentTarget));
                
                modals.empresas.addEventListener('change', (e) => {
                    if (e.target.matches('input[name="empresa_access"]')) {
                        const row = e.target.closest('tr');
                        const radio = row.querySelector('input[type="radio"]');
                        radio.disabled = !e.target.checked;
                        if (radio.disabled && radio.checked) radio.checked = false;
                    }
                });
            }

            // Botão voltar ao menu
            const backButton = document.getElementById('backToMenuFromUsersButton');
            if (backButton) {
                backButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            // ===============================================
            // INICIALIZAÇÃO
            // ===============================================
            
            // Carrega dados reais do backend na inicialização
            loadUsersFromBackend();
            loadEmpresasFromBackend();
        }

        // ===============================================
        // LÓGICA DA PÁGINA DE NOVO PEDIDO
        // ===============================================
        // FUNÇÃO PARA CONFIGURAR MODO DE EDIÇÃO
        // ===============================================
        function setupModoEdicao(dadosEdicao) {
            console.log('🔧 Configurando modo de edição:', dadosEdicao);
            
            // Alterar título da página para indicar modo de edição
            const titleElement = document.getElementById('page-title');
            const subtitleElement = document.getElementById('page-subtitle');
            
            if (titleElement) {
                titleElement.innerHTML = `
                    <i class="fas fa-edit mr-3 text-orange-500"></i>
                    Editar Pedido #${dadosEdicao.numeroDoPedido}
                    <span class="ml-3 px-3 py-1 bg-orange-100 text-orange-700 text-sm rounded-full">Modo Edição</span>
                `;
                titleElement.className = 'text-3xl font-bold text-orange-600 flex items-center';
            }
            
            if (subtitleElement) {
                subtitleElement.textContent = 'Você tem 1 hora para editar este pedido após a criação.';
                subtitleElement.className = 'text-orange-500';
            }
            
            // Alterar texto dos botões
            const salvarButton = document.getElementById('salvarPedidoButton');
            const salvarRascunhoButton = document.getElementById('salvarRascunhoButton');
            
            if (salvarButton) {
                salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Alterações';
                salvarButton.className = salvarButton.className.replace('bg-green-600', 'bg-orange-600').replace('hover:bg-green-700', 'hover:bg-orange-700');
            }
            
            if (salvarRascunhoButton) {
                salvarRascunhoButton.style.display = 'none'; // Esconder botão de rascunho em modo edição
            }
            
            // Adicionar aviso sobre tempo de edição
            adicionarAvisoTempoEdicao(dadosEdicao);
        }
        
        function preencherCamposEdicao(dados) {
            console.log('📝 Preenchendo campos para edição:', dados);
            
            // Preencher número do pedido
            if (dados.numeroDoPedido) {
                const numeroPedidoElement = document.getElementById('numeroPedido');
                if (numeroPedidoElement) {
                    console.log('✅ Preenchendo número do pedido:', dados.numeroDoPedido);
                    numeroPedidoElement.value = dados.numeroDoPedido;
                } else {
                    console.warn('⚠️ Campo numeroPedido não encontrado');
                }
            }
            
            // Preencher data
            if (dados.data) {
                const dataPedidoElement = document.getElementById('dataPedido');
                if (dataPedidoElement) {
                    console.log('✅ Preenchendo data:', dados.data);
                    dataPedidoElement.value = dados.data;
                } else {
                    console.warn('⚠️ Campo dataPedido não encontrado');
                }
            }
            
            // Preencher observações
            if (dados.observacoes) {
                const observacoesElement = document.getElementById('observacoesPedido');
                if (observacoesElement) {
                    console.log('✅ Preenchendo observações:', dados.observacoes);
                    observacoesElement.value = dados.observacoes;
                } else {
                    console.warn('⚠️ Campo observacoesPedido não encontrado');
                }
            }
            
            // Preencher fornecedor
            if (dados.cnpjFornecedor || dados.fornecedor) {
                const fornecedorElement = document.getElementById('fornecedorPedido');
                if (fornecedorElement) {
                    // Aguardar o carregamento dos fornecedores antes de selecionar
                    setTimeout(() => {
                        const cnpjProcurado = dados.cnpjFornecedor;
                        const nomeProcurado = dados.fornecedor;
                        
                        console.log('🔍 Procurando fornecedor:', { cnpjProcurado, nomeProcurado });
                        
                        // Primeiro tentar por CNPJ
                        if (cnpjProcurado) {
                            const opcoes = fornecedorElement.options;
                            for (let i = 0; i < opcoes.length; i++) {
                                const opcao = opcoes[i];
                                if (opcao.dataset.cnpj === cnpjProcurado.toString()) {
                                    fornecedorElement.selectedIndex = i;
                                    console.log('✅ Fornecedor encontrado por CNPJ:', opcao.textContent);
                                    fornecedorElement.dispatchEvent(new Event('change'));
                                    return;
                                }
                            }
                        }
                        
                        // Se não encontrou por CNPJ, tentar por nome/razão
                        if (nomeProcurado) {
                            fornecedorElement.value = nomeProcurado;
                            if (fornecedorElement.selectedIndex > 0) {
                                console.log('✅ Fornecedor encontrado por nome:', nomeProcurado);
                                fornecedorElement.dispatchEvent(new Event('change'));
                                return;
                            }
                        }
                        
                        console.warn('⚠️ Fornecedor não encontrado nos selects disponíveis');
                    }, 1500);
                } else {
                    console.warn('⚠️ Campo fornecedorPedido não encontrado');
                }
            }
            
            // Preencher veículo
            if (dados.nomeVeiculo) {
                const veiculoElement = document.getElementById('nomeVeiculo');
                if (veiculoElement) {
                    setTimeout(() => {
                        console.log('✅ Preenchendo veículo:', dados.nomeVeiculo);
                        veiculoElement.value = dados.nomeVeiculo;
                    }, 1000);
                } else {
                    console.warn('⚠️ Campo nomeVeiculo não encontrado');
                }
            }
            
            // Preencher placa
            if (dados.placaVeiculo) {
                const placaElement = document.getElementById('placaVeiculo');
                if (placaElement) {
                    console.log('✅ Preenchendo placa:', dados.placaVeiculo);
                    placaElement.value = dados.placaVeiculo;
                } else {
                    console.warn('⚠️ Campo placaVeiculo não encontrado');
                }
            }
            
            // Preencher itens
            if (dados.itens && dados.itens.length > 0) {
                console.log('📋 Preenchendo itens:', dados.itens.length, 'itens encontrados');
                window.itensAdicionados = dados.itens.map(item => ({
                    descricao: item.descricao,
                    quantidade: parseFloat(item.quantidade) || 0,
                    unidade: item.unidade || 'UN',
                    precoUnitario: parseFloat(item.precoUnitario) || 0,
                    subtotal: parseFloat(item.totalItem) || (parseFloat(item.quantidade) * parseFloat(item.precoUnitario))
                }));
                
                setTimeout(() => {
                    console.log('📊 Carregando itens na tabela...');
                    preencherItensContainer(dados.itens);
                    
                    // Atualizar total geral
                    const totalCalculado = dados.totalGeral || dados.itens.reduce((total, item) => {
                        return total + (parseFloat(item.totalItem) || 0);
                    }, 0);
                    
                    console.log('💰 Total calculado:', totalCalculado);
                    
                    const totalGeralElement = document.getElementById('totalGeral');
                    const summaryTotalElement = document.getElementById('summary-total');
                    
                    if (totalGeralElement) {
                        totalGeralElement.textContent = `R$ ${totalCalculado.toFixed(2).replace('.', ',')}`;
                        console.log('✅ Total geral atualizado no campo totalGeral');
                    }
                    
                    if (summaryTotalElement) {
                        summaryTotalElement.textContent = `R$ ${totalCalculado.toFixed(2).replace('.', ',')}`;
                        console.log('✅ Total geral atualizado no resumo');
                    }
                    
                }, 1500);
            }
        }
        
        function adicionarAvisoTempoEdicao(dados) {
            // Verificar se ainda pode editar
            const podeEditar = verificarSePermiteEdicao(dados);
            
            if (!podeEditar) {
                // Se não pode mais editar, mostrar aviso e desabilitar campos
                const container = document.querySelector('.max-w-7xl');
                if (container) {
                    const aviso = document.createElement('div');
                    aviso.className = 'bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-6';
                    aviso.innerHTML = `
                        <div class="flex items-center">
                            <i class="fas fa-exclamation-triangle mr-2"></i>
                            <strong>Tempo de edição expirado!</strong>
                            <span class="ml-2">Este pedido só pode ser editado dentro de 1 hora após a criação.</span>
                        </div>
                    `;
                    container.insertBefore(aviso, container.firstChild);
                }
                
                // Desabilitar todos os campos
                const campos = document.querySelectorAll('input, select, textarea, button');
                campos.forEach(campo => {
                    if (campo.id !== 'backToMenuFromPedidoButton') {
                        campo.disabled = true;
                    }
                });
                
                return false;
            } else {
                // Mostrar tempo restante
                const container = document.querySelector('.max-w-7xl');
                if (container && dados.dataHoraCriacao) {
                    const agora = new Date();
                    const criacao = new Date(dados.dataHoraCriacao);
                    const diferencaMinutos = Math.ceil(60 - ((agora - criacao) / (1000 * 60)));
                    
                    const aviso = document.createElement('div');
                    aviso.className = 'bg-yellow-100 border border-yellow-400 text-yellow-700 px-4 py-3 rounded mb-6';
                    aviso.innerHTML = `
                        <div class="flex items-center">
                            <i class="fas fa-clock mr-2"></i>
                            <strong>Modo Edição:</strong>
                            <span class="ml-2">Você tem aproximadamente ${diferencaMinutos} minutos restantes para editar este pedido.</span>
                        </div>
                    `;
                    container.insertBefore(aviso, container.firstChild);
                }
            }
            
            return podeEditar;
        }
        
        // ===============================================
        let itemCounter = 0;
        
        // Cache para fornecedores e veículos
        let cacheData = {
            fornecedores: null,
            veiculos: null,
            timestamp: null,
            expireTime: 5 * 60 * 1000 // 5 minutos
        };
        
        function isCacheValid() {
            return cacheData.timestamp && 
                   cacheData.fornecedores && 
                   cacheData.veiculos && 
                   (new Date().getTime() - cacheData.timestamp) < cacheData.expireTime;
        }
        
        function setupPedidoScreen(modoEdicao = false, dadosEdicao = null) {
            console.log('🔧 setupPedidoScreen iniciado', { modoEdicao, dadosEdicao });
            
            // Verificar se está em modo visualização (diferente de edição)
            const modoVisualizacao = localStorage.getItem('modoVisualizacao');
            const origemVisualizacao = localStorage.getItem('origemVisualizacao'); // 'dashboard-admin' ou 'menu-usuario'
            
            console.log('🔧 modoVisualizacao:', modoVisualizacao);
            console.log('🔧 origemVisualizacao:', origemVisualizacao);
            console.log('🔧 modoEdicao:', modoEdicao);
            
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            const idDaEmpresa = empresaAtual.id || empresaAtual.codigo;

            if (!idDaEmpresa) {
                showGlobalModal('Erro', 'Não foi possível identificar a empresa. Por favor, faça o login novamente.');
                return;
            }
            
            // Configurar interface para modo de edição
            if (modoEdicao && dadosEdicao) {
                console.log('🔧 Configurando modo de edição:', dadosEdicao);
                // Configurar campos primeiro, depois carregar selects
                setupModoEdicao(dadosEdicao);
            }
            
            // Carregar fornecedores e veículos com cache e em paralelo
            console.log('🔧 Carregando fornecedores e veículos...');
            
            let fornecedoresCarregados = false;
            let veiculosCarregados = false;
            
            function verificarCarregamentoCompleto() {
                if (fornecedoresCarregados && veiculosCarregados && modoEdicao && dadosEdicao) {
                    console.log('🔧 Todos os dados carregados, preenchendo campos...');
                    setTimeout(() => {
                        preencherCamposEdicao(dadosEdicao);
                    }, 300); // Reduzido de 500ms para 300ms
                }
            }
            
            // Verificar se podemos usar o cache
            if (isCacheValid()) {
                console.log('🚀 Usando dados do cache para carregamento rápido');
                populateFornecedores(cacheData.fornecedores);
                populateVeiculos(cacheData.veiculos);
                fornecedoresCarregados = true;
                veiculosCarregados = true;
                verificarCarregamentoCompleto();
                return;
            }
            
            // Carregamento paralelo dos dados
            console.log('⏱️ Iniciando carregamento paralelo...');
            const startTime = performance.now();
            
            // Mostrar indicador de carregamento
            const loadingIndicator = document.createElement('div');
            loadingIndicator.id = 'loading-indicator';
            loadingIndicator.className = 'fixed top-4 right-4 bg-blue-500 text-white px-4 py-2 rounded-lg shadow-lg z-50 flex items-center';
            loadingIndicator.innerHTML = `
                <div class="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                Carregando dados...
            `;
            document.body.appendChild(loadingIndicator);
            
            // Função para popular fornecedores
            function populateFornecedores(fornecedores) {
                const selectFornecedor = document.getElementById('fornecedorPedido');
                if (selectFornecedor) {
                    selectFornecedor.innerHTML = '<option value="">Selecione um Fornecedor</option>';
                    fornecedores.forEach(f => {
                        const option = document.createElement('option');
                        option.value = f.razao || f.codigo;
                        option.textContent = f.razao || f.codigo;
                        if (f.grupo) option.dataset.grupo = f.grupo;
                        if (f.estado) option.dataset.estado = f.estado;
                        if (f.cnpj) option.dataset.cnpj = f.cnpj;
                        selectFornecedor.appendChild(option);
                    });
                    
                    // Configurar event listener para atualizar o estado quando fornecedor for selecionado
                    selectFornecedor.addEventListener('change', function() {
                        const selectedOption = this.options[this.selectedIndex];
                        const estadoElement = document.getElementById('fornecedor-estado');
                        
                        if (estadoElement) {
                            if (selectedOption && selectedOption.dataset.estado) {
                                estadoElement.textContent = selectedOption.dataset.estado;
                                console.log('🌍 Estado atualizado para:', selectedOption.dataset.estado);
                            } else {
                                estadoElement.textContent = 'Selecione';
                                console.log('🌍 Estado resetado');
                            }
                        }
                    });
                }
            }
            
            // Função para popular veículos
            function populateVeiculos(veiculos) {
                const selectVeiculo = document.getElementById('nomeVeiculo');
                if (selectVeiculo) {
                    selectVeiculo.innerHTML = '<option value="">Selecione um veículo</option>';
                    
                    veiculos.forEach((v, index) => {
                        // Tentar diferentes propriedades para encontrar o nome
                        let nomeVeiculo = null;
                        
                        if (typeof v === 'string' && v.trim() !== '' && v !== 'undefined') {
                            nomeVeiculo = v.trim();
                        } else if (typeof v === 'object' && v !== null) {
                            nomeVeiculo = v.nomeVeiculo || v.nome || v.veiculo || v.descricao;
                        }
                        
                        if (nomeVeiculo && nomeVeiculo.trim() !== '' && nomeVeiculo !== 'undefined') {
                            const option = document.createElement('option');
                            option.value = nomeVeiculo.trim();
                            option.textContent = nomeVeiculo.trim();
                            selectVeiculo.appendChild(option);
                        } else {
                            console.warn(`⚠️ Veículo inválido no índice ${index}:`, v);
                        }
                    });
                }
            }
            
            // Carregamento paralelo usando Promise.all
            Promise.all([
                // Carregar fornecedores
                new Promise((resolve, reject) => {
                    google.script.run.withSuccessHandler(fornecedores => {
                        try {
                            populateFornecedores(fornecedores);
                            cacheData.fornecedores = fornecedores;
                            console.log('✅ Fornecedores carregados:', fornecedores.length);
                            fornecedoresCarregados = true;
                            resolve(fornecedores);
                        } catch (error) {
                            reject(error);
                        }
                    }).withFailureHandler(reject).getFornecedoresList();
                }),
                
                // Carregar veículos
                new Promise((resolve, reject) => {
                    google.script.run.withSuccessHandler(veiculos => {
                        try {
                            populateVeiculos(veiculos);
                            cacheData.veiculos = veiculos;
                            console.log('✅ Veículos carregados:', veiculos.length);
                            veiculosCarregados = true;
                            resolve(veiculos);
                        } catch (error) {
                            reject(error);
                        }
                    }).withFailureHandler(reject).getVeiculosList();
                })
            ]).then(() => {
                // Atualizar timestamp do cache
                cacheData.timestamp = new Date().getTime();
                
                const endTime = performance.now();
                console.log(`🚀 Carregamento paralelo concluído em ${Math.round(endTime - startTime)}ms`);
                
                // Remover indicador de carregamento
                const indicator = document.getElementById('loading-indicator');
                if (indicator) {
                    indicator.remove();
                }
                
                verificarCarregamentoCompleto();
            }).catch(error => {
                console.error('❌ Erro no carregamento paralelo:', error);
                
                // Remover indicador de carregamento mesmo em caso de erro
                const indicator = document.getElementById('loading-indicator');
                if (indicator) {
                    indicator.remove();
                }
            });
            
            // Verificar se estamos em modo visualização
            if (modoVisualizacao === 'true') {
                console.log("🔍 Modo visualização detectado");
                
                // Verificar se já temos dados do pedido
                const pedidoParaVisualizar = localStorage.getItem('pedidoParaVisualizar');
                
                if (pedidoParaVisualizar) {
                    console.log("📋 Dados do pedido já disponíveis, prosseguindo com visualização");
                    
                    // Detectar contexto e usar a função apropriada
                    if (origemVisualizacao === 'dashboard-admin') {
                        console.log("🔍 Modo visualização: Dashboard Admin - NÃO chamando setupModoVisualizacao");
                        // Não chamar setupModoVisualizacao aqui, será chamado pelo finalizarPreenchimento
                    } else {
                        console.log("🔍 Modo visualização: Menu Usuário - chamando setupModoVisualizacao");
                        setupModoVisualizacao();
                    }
                    
                    setTimeout(() => {
                        console.log('🔧 Chamando preencherTelaPedidoExistente após 1000ms');
                        preencherTelaPedidoExistente();
                    }, 1000);
                    return;
                } else {
                    console.log("⏳ Aguardando dados do pedido serem carregados...");
                    // Aguardar um pouco e tentar novamente (dados podem estar sendo buscados)
                    setTimeout(() => {
                        const pedidoTardio = localStorage.getItem('pedidoParaVisualizar');
                        if (pedidoTardio) {
                            console.log("📋 Dados do pedido carregados tardiamente, prosseguindo");
                            if (origemVisualizacao === 'dashboard-admin') {
                                console.log("🔍 Modo visualização: Dashboard Admin - NÃO chamando setupModoVisualizacao");
                            } else {
                                console.log("🔍 Modo visualização: Menu Usuário - chamando setupModoVisualizacao");
                                setupModoVisualizacao();
                            }
                            preencherTelaPedidoExistente();
                        } else {
                            console.warn("⚠️ Modo visualização ativo mas dados do pedido não encontrados");
                            // Limpar modo visualização e continuar como criação
                            localStorage.removeItem('modoVisualizacao');
                            localStorage.removeItem('origemVisualizacao');
                        }
                    }, 1000);
                    return;
                }
            }

            // Modo criação - configuração normal
            const salvarPedidoBtn = document.getElementById('salvarPedidoButton');
            const salvarRascunhoBtn = document.getElementById('salvarRascunhoButton');
            const cancelarPedidoBtn = document.getElementById('cancelarPedidoButton');
            const addItemBtn = document.getElementById('addItemButton');
            
            console.log('🔧 Configurando event listeners para modo criação');
            console.log('🔧 Botões encontrados:', {
                salvarPedido: !!salvarPedidoBtn,
                salvarRascunho: !!salvarRascunhoBtn,
                cancelar: !!cancelarPedidoBtn,
                addItem: !!addItemBtn
            });
            
            if (salvarPedidoBtn) {
                // Verificar se já tem event listener para evitar duplicação
                const newButton = salvarPedidoBtn.cloneNode(true);
                salvarPedidoBtn.parentNode.replaceChild(newButton, salvarPedidoBtn);
                newButton.addEventListener('click', salvarPedido);
                console.log('✅ Event listener adicionado ao botão salvar pedido');
            } else {
                console.warn('⚠️ Botão salvarPedidoButton não encontrado');
            }
            if (salvarRascunhoBtn) {
                const newRascunhoButton = salvarRascunhoBtn.cloneNode(true);
                salvarRascunhoBtn.parentNode.replaceChild(newRascunhoButton, salvarRascunhoBtn);
                newRascunhoButton.addEventListener('click', salvarRascunho);
                console.log('✅ Event listener adicionado ao botão salvar rascunho');
            } else {
                console.warn('⚠️ Botão salvarRascunhoButton não encontrado');
            }
            
            if (cancelarPedidoBtn) {
                const newCancelarButton = cancelarPedidoBtn.cloneNode(true);
                cancelarPedidoBtn.parentNode.replaceChild(newCancelarButton, cancelarPedidoBtn);
                newCancelarButton.addEventListener('click', () => {
                    // Limpar dados de edição de rascunho
                    limparDadosEdicaoRascunho();
                    loadPageContent('Menu', setupMenuScreen);
                });
                console.log('✅ Event listener adicionado ao botão cancelar');
            } else {
                console.warn('⚠️ Botão cancelarPedidoButton não encontrado');
            }
            if (addItemBtn) {
                addItemBtn.addEventListener('click', adicionarItem);
                addItemBtn.hasClickListener = true; // Marcar como configurado
                console.log('✅ Event listener adicionado ao botão Adicionar Item');
            } else {
                console.warn('⚠️ Botão addItemButton não encontrado!');
            }
            
            // Configurar eventos do formulário de entrada principal
            const itemQuantidadeInput = document.getElementById('itemQuantidade');
            const itemPrecoUnitarioInput = document.getElementById('itemPrecoUnitario');
            const itemDescricaoInput = document.getElementById('itemDescricao');
            const itemUnidadeSelect = document.getElementById('itemUnidade');
            
            if (itemQuantidadeInput && itemPrecoUnitarioInput) {
                // Calcular subtotal em tempo real
                const calcularSubtotal = () => {
                    const quantidade = parseFloat(itemQuantidadeInput.value) || 0;
                    const preco = parseFloat(itemPrecoUnitarioInput.value.replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
                    const subtotal = quantidade * preco;
                    
                    const liveSubtotalElement = document.getElementById('live-subtotal');
                    if (liveSubtotalElement) {
                        liveSubtotalElement.textContent = `R$ ${subtotal.toLocaleString('pt-BR', {minimumFractionDigits: 2})}`;
                    }
                };
                
                itemQuantidadeInput.addEventListener('input', calcularSubtotal);
                itemPrecoUnitarioInput.addEventListener('input', calcularSubtotal);
            }
            
            // Configurar Enter key no formulário principal
            if (itemDescricaoInput && itemQuantidadeInput && itemUnidadeSelect && itemPrecoUnitarioInput) {
                const campos = [itemDescricaoInput, itemQuantidadeInput, itemUnidadeSelect, itemPrecoUnitarioInput];
                
                campos.forEach(campo => {
                    campo.addEventListener('keydown', (e) => {
                        if (e.key === 'Enter') {
                            e.preventDefault();
                            
                            // Validar se todos os campos estão preenchidos
                            const descricao = itemDescricaoInput.value.trim();
                            const quantidade = parseFloat(itemQuantidadeInput.value) || 0;
                            const unidade = itemUnidadeSelect.value;
                            const preco = parseFloat(itemPrecoUnitarioInput.value.replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
                            
                            if (descricao && quantidade > 0 && unidade && preco > 0) {
                                adicionarItem();
                            } else {
                                // Focar no próximo campo não preenchido
                                if (!descricao) {
                                    itemDescricaoInput.focus();
                                } else if (quantidade <= 0) {
                                    itemQuantidadeInput.focus();
                                } else if (!unidade) {
                                    itemUnidadeSelect.focus();
                                } else if (preco <= 0) {
                                    itemPrecoUnitarioInput.focus();
                                }
                            }
                        }
                    });
                });
            }
            
            const itensContainer = document.getElementById('itensContainer');
            if (itensContainer) {
                console.log('ℹ️ Container de itens encontrado (usado para compatibilidade)');
            } else {
                console.warn('⚠️ Container de itens não encontrado!');
            }
            
            // Inicializar variáveis globais
            window.itensAdicionados = [];
            window.itemCounter = 0;
            
            const today = new Date();
            const dataPedidoElement = document.getElementById('dataPedido');
            if (dataPedidoElement) {
                // Usar data local ao invés de UTC para evitar problemas de fuso horário
                const ano = today.getFullYear();
                const mes = String(today.getMonth() + 1).padStart(2, '0');
                const dia = String(today.getDate()).padStart(2, '0');
                const dataLocal = `${ano}-${mes}-${dia}`;
                
                dataPedidoElement.value = dataLocal;
                console.log('📅 Data do pedido inicializada:', dataLocal);
            }

            // Garantir que os eventos sejam configurados mesmo com atraso no carregamento
            setTimeout(() => {
                console.log('🔄 Verificação adicional dos event listeners...');
                
                const addItemBtnCheck = document.getElementById('addItemButton');
                if (addItemBtnCheck && !addItemBtnCheck.hasClickListener) {
                    console.log('🔧 Configurando event listener do botão (segunda tentativa)');
                    addItemBtnCheck.addEventListener('click', adicionarItem);
                    addItemBtnCheck.hasClickListener = true;
                }
            }, 500);
        }

        // Função específica para garantir modo somente leitura
        function garantirModoSomenteGeitura() {
            console.log('🔒 Função garantirModoSomenteGeitura executada');
            
            // Tentar múltiplos seletores para encontrar o container
            let container = document.querySelector('.max-w-7xl.mx-auto');
            if (!container) container = document.querySelector('main');
            if (!container) container = document.querySelector('body');
            if (!container) container = document;
            
            console.log('🔍 Container usado:', container);
            
            if (container) {
                // Buscar de forma mais específica
                const inputs = container.querySelectorAll('input');
                const selects = container.querySelectorAll('select'); 
                const textareas = container.querySelectorAll('textarea');
                const buttons = container.querySelectorAll('button');
                
                console.log(`� Elementos encontrados: ${inputs.length} inputs, ${selects.length} selects, ${textareas.length} textareas, ${buttons.length} buttons`);
                
                // Desabilitar inputs
                inputs.forEach(input => {
                    if (!input.readOnly) {
                        input.disabled = true;
                        input.readOnly = true;
                        input.classList.add('bg-gray-100', 'text-gray-600', 'cursor-not-allowed');
                        console.log(`🔒 Input desabilitado: ${input.id || input.name || input.type}`);
                    }
                });
                
                // Desabilitar selects
                selects.forEach(select => {
                    select.disabled = true;
                    select.classList.add('bg-gray-100', 'text-gray-600', 'cursor-not-allowed');
                    console.log(`🔒 Select desabilitado: ${select.id || select.name}`);
                });
                
                // Desabilitar textareas
                textareas.forEach(textarea => {
                    textarea.disabled = true;
                    textarea.readOnly = true;
                    textarea.classList.add('bg-gray-100', 'text-gray-600', 'cursor-not-allowed');
                    console.log(`🔒 Textarea desabilitado: ${textarea.id || textarea.name}`);
                });
                
                // Desabilitar botões (exceto navegação)
                buttons.forEach(button => {
                    const id = button.id || '';
                    const isNavButton = id.includes('voltar') || id.includes('imprimir') || id.includes('Menu') || id.includes('Dashboard');
                    
                    if (!isNavButton) {
                        button.disabled = true;
                        button.classList.add('opacity-50', 'cursor-not-allowed');
                        console.log(`🔒 Botão desabilitado: ${button.id || button.textContent?.trim().substring(0, 20)}`);
                    } else {
                        console.log(`✅ Botão preservado: ${button.id}`);
                    }
                });
                
                console.log('✅ Modo somente leitura garantido');
            } else {
                console.error('❌ Nenhum container encontrado!');
            }
        }

        function setupModoVisualizacao() {
            console.log("🔍 setupModoVisualizacao - Configurando modo visualização de pedido");
            
            // Trocar os botões da área de ações
            console.log('🔍 Procurando área de botões...');
            let botoesArea = document.querySelector('.pt-6.space-y-3');
            console.log('🔍 Tentativa 1 (.pt-6.space-y-3):', botoesArea);
            
            if (!botoesArea) {
                // Tentar seletores alternativos
                botoesArea = document.querySelector('[class*="space-y-3"]');
                console.log('🔍 Tentativa 2 ([class*="space-y-3"]):', botoesArea);
            }
            
            if (!botoesArea) {
                botoesArea = document.querySelector('#salvarPedidoButton')?.parentElement;
                console.log('🔍 Tentativa 3 (parent do salvarPedidoButton):', botoesArea);
            }
            
            if (!botoesArea) {
                botoesArea = document.querySelector('#salvarRascunhoButton')?.parentElement;
                console.log('🔍 Tentativa 4 (parent do salvarRascunhoButton):', botoesArea);
            }
            
            if (botoesArea) {
                console.log('🔧 setupModoVisualizacao - Área de botões encontrada, substituindo');
                const numeroPedido = localStorage.getItem('numeroPedidoVisualizacao') || '';
                
                botoesArea.innerHTML = `
                    <button type="button" id="voltarMenuButton" class="w-full bg-gray-600 text-white font-bold py-3 px-4 rounded-lg shadow-md hover:bg-gray-700">
                        <i class="fas fa-arrow-left mr-2"></i>Voltar ao Menu
                    </button>
                    <button type="button" id="imprimirPedidoButton" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg shadow-md hover:bg-blue-700">
                        <i class="fas fa-print mr-2"></i>Imprimir Pedido
                    </button>
                `;
                
                console.log('🔧 setupModoVisualizacao - Botões substituídos, adicionando event listeners');
                
                // Adicionar eventos aos novos botões
                document.getElementById('voltarMenuButton').addEventListener('click', () => {
                    console.log('🔙 Botão Voltar ao Menu clicado');
                    localStorage.removeItem('modoVisualizacao');
                    localStorage.removeItem('numeroPedidoVisualizacao');
                    localStorage.removeItem('origemVisualizacao');
                    loadPageContent('Menu', setupMenuScreen);
                });
                
                document.getElementById('imprimirPedidoButton').addEventListener('click', () => {
                    console.log('🖨️ Botão Imprimir clicado para pedido:', numeroPedido);
                    abrirImpressaoPedido(numeroPedido);
                });
                
                console.log('✅ setupModoVisualizacao - Event listeners adicionados');
            } else {
                console.warn('⚠️ setupModoVisualizacao - Área de botões não encontrada');
            }
            
            // Esconder o botão "Adicionar Item" 
            const addItemButton = document.getElementById('addItemButton');
            if (addItemButton) {
                addItemButton.style.display = 'none';
            }
            
            // Desabilitar todos os campos para modo somente leitura
            setTimeout(() => {
                console.log('🔒 Iniciando desabilitação de campos...');
                
                // Tentar múltiplos seletores para encontrar o container
                let container = document.querySelector('.max-w-7xl.mx-auto');
                if (!container) container = document.querySelector('main');
                if (!container) container = document.querySelector('body');
                
                console.log('🔍 Container encontrado:', container ? 'SIM' : 'NÃO');
                
                if (container) {
                    // Buscar campos de forma mais específica
                    const inputs = container.querySelectorAll('input:not([readonly])');
                    const selects = container.querySelectorAll('select');
                    const textareas = container.querySelectorAll('textarea');
                    
                    console.log(`🔒 Encontrados ${inputs.length} inputs, ${selects.length} selects, ${textareas.length} textareas`);
                    
                    [...inputs, ...selects, ...textareas].forEach(campo => {
                        campo.disabled = true;
                        campo.readOnly = true; // Adicionar readonly também
                        campo.classList.add('bg-gray-100', 'text-gray-600', 'cursor-not-allowed');
                        console.log(`🔒 Campo desabilitado: ${campo.id || campo.name || campo.tagName}`);
                    });
                    
                    console.log('✅ Todos os campos foram desabilitados para modo visualização');
                    
                    // Esconder seção de adicionar itens
                    const secoes = container.querySelectorAll('.bg-white');
                    secoes.forEach(secao => {
                        const titulo = secao.querySelector('h3');
                        if (titulo && titulo.textContent.includes('Adicionar Itens')) {
                            secao.style.display = 'none';
                            console.log('✅ Seção "Adicionar Itens" escondida');
                        }
                    });
                } else {
                    console.warn('⚠️ Container principal não encontrado para desabilitar campos');
                }
            }, 1500); // Aumentar timeout para executar após preenchimento completo
            
            // Alterar o título da tela
            const titulo = document.querySelector('h1');
            if (titulo) {
                titulo.textContent = 'Visualizar Pedido de Compra';
                titulo.classList.add('text-blue-800');
            }
            
            // Garantir modo somente leitura com função específica
            setTimeout(() => {
                garantirModoSomenteGeitura();
            }, 2000);
        }

        function handleItemInputChange(event) {
            if (event.target.name === 'itemQuantidade' || event.target.name === 'itemPrecoUnitario') {
                const itemRow = event.target.closest('.item-row');
                if (itemRow) {
                    const itemId = itemRow.dataset.itemId;
                    calcularTotalItem(itemId);
                }
            }
        }

        function handleItemButtonClick(event) {
            // Função mantida para compatibilidade com layout antigo
            // Nova implementação usa removerItem() diretamente
            if (event.target.classList.contains('remove-item-btn')) {
                const itemRow = event.target.closest('.item-row');
                if (document.querySelectorAll('.item-row').length > 1) {
                    itemRow.remove();
                    calcularTotalGeral();
                } else {
                    showGlobalModal('Aviso', 'Não é possível remover o último item do pedido.');
                }
            }
        }

        function handleItemKeyDown(event) {
            console.log('⌨️ Tecla pressionada:', event.key, 'Target:', event.target);
            
            // Verificar se a tecla pressionada é Enter
            if (event.key === 'Enter') {
                console.log('↵ Enter detectado!');
                
                // Verificar se o elemento focado está dentro de um item-row
                const itemRow = event.target.closest('.item-row');
                if (itemRow) {
                    console.log('📦 Item row encontrado:', itemRow.dataset.itemId);
                    
                    // Prevenir comportamento padrão do Enter
                    event.preventDefault();
                    event.stopPropagation();
                    
                    // Verificar se todos os campos obrigatórios do item atual estão preenchidos
                    const descricao = itemRow.querySelector('input[name="itemDescricao"]');
                    const quantidade = itemRow.querySelector('input[name="itemQuantidade"]');
                    const unidade = itemRow.querySelector('select[name="itemUnidade"]');
                    const precoUnitario = itemRow.querySelector('input[name="itemPrecoUnitario"]');
                    
                    // Log dos valores encontrados
                    console.log('🔍 Validação dos campos:');
                    console.log('  - Descrição:', descricao?.value?.trim() || 'vazio');
                    console.log('  - Quantidade:', quantidade?.value || 'vazio');
                    console.log('  - Unidade:', unidade?.value || 'vazio');
                    console.log('  - Preço:', precoUnitario?.value || 'vazio');
                    
                    // Validar se os campos essenciais estão preenchidos
                    if (descricao && descricao.value.trim() && 
                        quantidade && quantidade.value && parseFloat(quantidade.value) > 0 &&
                        unidade && unidade.value && 
                        precoUnitario && precoUnitario.value && parseFloat(precoUnitario.value) > 0) {
                        
                        console.log('✅ Todos os campos preenchidos - adicionando novo item');
                        
                        // Adicionar novo item
                        adicionarItem();
                        
                        // Focar no campo de descrição do novo item após um pequeno delay
                        setTimeout(() => {
                            const novoItem = document.querySelector('.item-row:last-child');
                            if (novoItem) {
                                const novaDescricao = novoItem.querySelector('input[name="itemDescricao"]');
                                if (novaDescricao) {
                                    novaDescricao.focus();
                                    console.log('🎯 Foco movido para o novo item');
                                }
                            }
                        }, 100);
                    } else {
                        console.log('⚠️ Campos obrigatórios não preenchidos');
                        // Mostrar toast informando quais campos precisam ser preenchidos
                        showToast('Preencha todos os campos obrigatórios (Descrição, Quantidade > 0, Unidade e Preço > 0) antes de adicionar um novo item.', 'warning', 'Campos Obrigatórios');
                        
                        // Focar no primeiro campo vazio
                        if (!descricao?.value?.trim()) {
                            descricao?.focus();
                        } else if (!quantidade?.value || parseFloat(quantidade.value) <= 0) {
                            quantidade?.focus();
                        } else if (!unidade?.value) {
                            unidade?.focus();
                        } else if (!precoUnitario?.value || parseFloat(precoUnitario.value) <= 0) {
                            precoUnitario?.focus();
                        }
                    }
                } else {
                    console.log('❌ Item row não encontrado para o elemento focado');
                }
            }
        }

        function adicionarItem() {
            console.log('🔧 adicionarItem() chamada');
            
            // Pegar valores dos campos do formulário
            const descricao = document.getElementById('itemDescricao').value.trim();
            const quantidade = parseFloat(document.getElementById('itemQuantidade').value) || 0;
            const unidade = document.getElementById('itemUnidade').value;
            const precoUnitario = parseFloat(document.getElementById('itemPrecoUnitario').value.replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
            
            // Validações
            if (!descricao) {
                showToast('Por favor, informe a descrição do item.', 'warning');
                document.getElementById('itemDescricao').focus();
                return;
            }
            
            if (quantidade <= 0) {
                showToast('A quantidade deve ser maior que zero.', 'warning');
                document.getElementById('itemQuantidade').focus();
                return;
            }
            
            if (!unidade) {
                showToast('Por favor, selecione a unidade.', 'warning');
                document.getElementById('itemUnidade').focus();
                return;
            }
            
            if (precoUnitario <= 0) {
                showToast('O preço unitário deve ser maior que zero.', 'warning');
                document.getElementById('itemPrecoUnitario').focus();
                return;
            }
            
            // Calcular subtotal
            const subtotal = quantidade * precoUnitario;
            
            // Incrementar contador
            itemCounter++;
            
            // Obter a tabela
            const tableBody = document.getElementById('itensTableBody');
            if (!tableBody) {
                console.error('❌ Tabela de itens não encontrada!');
                return;
            }
            
            // Remover linha "Nenhum item adicionado" se existir
            const noItemsRow = document.getElementById('no-items-row');
            if (noItemsRow) {
                noItemsRow.remove();
            }
            
            // Criar nova linha na tabela
            const row = document.createElement('tr');
            row.setAttribute('data-item-id', itemCounter);
            row.className = 'hover:bg-gray-50';
            row.innerHTML = `
                <td class="px-6 py-4 text-sm text-gray-900">${descricao}</td>
                <td class="px-6 py-4 text-sm text-gray-900 text-center">${quantidade.toLocaleString('pt-BR')}</td>
                <td class="px-6 py-4 text-sm text-gray-900 text-center">${unidade}</td>
                <td class="px-6 py-4 text-sm text-gray-900 text-right">R$ ${precoUnitario.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                <td class="px-6 py-4 text-sm font-medium text-gray-900 text-right">R$ ${subtotal.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                <td class="px-6 py-4 text-center">
                    <button type="button" class="text-red-600 hover:text-red-800 font-medium" onclick="removerItem(${itemCounter})">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            `;
            
            // Adicionar à tabela
            tableBody.appendChild(row);
            
            // Armazenar dados do item (para salvar posteriormente)
            if (!window.itensAdicionados) {
                window.itensAdicionados = [];
            }
            window.itensAdicionados.push({
                id: itemCounter,
                descricao: descricao,
                quantidade: quantidade,
                unidade: unidade,
                precoUnitario: precoUnitario,
                subtotal: subtotal
            });
            
            // Limpar formulário
            document.getElementById('itemDescricao').value = '';
            document.getElementById('itemQuantidade').value = '1';
            document.getElementById('itemUnidade').value = 'UN'; // Voltar para UN como padrão
            document.getElementById('itemPrecoUnitario').value = '';
            document.getElementById('live-subtotal').textContent = 'R$ 0,00';
            
            // Recalcular totais
            calcularTotalGeral();
            
            // Focar no campo descrição para próximo item
            document.getElementById('itemDescricao').focus();
            
            console.log(`✅ Item ${itemCounter} adicionado à tabela com sucesso`);
        }

        function removerItem(itemId) {
            console.log('🗑️ Removendo item:', itemId);
            
            // Remover da tabela
            const row = document.querySelector(`tr[data-item-id="${itemId}"]`);
            if (row) {
                row.remove();
            }
            
            // Remover do array de itens
            if (window.itensAdicionados) {
                window.itensAdicionados = window.itensAdicionados.filter(item => item.id !== itemId);
            }
            
            // Verificar se não há mais itens e mostrar mensagem
            const tableBody = document.getElementById('itensTableBody');
            if (tableBody && tableBody.children.length === 0) {
                const noItemsRow = document.createElement('tr');
                noItemsRow.id = 'no-items-row';
                noItemsRow.innerHTML = '<td colspan="6" class="text-center text-gray-500 py-10">Nenhum item adicionado.</td>';
                tableBody.appendChild(noItemsRow);
            }
            
            // Recalcular totais
            calcularTotalGeral();
            
            console.log('✅ Item removido com sucesso');
        }

        // Função de debug para testar o botão adicionar item
        function testarBotaoAdicionarItem() {
            console.log('🧪 TESTE: Verificando botão Adicionar Item');
            const botao = document.getElementById('addItemButton');
            console.log('Botão encontrado:', !!botao);
            if (botao) {
                console.log('Botão visível:', botao.offsetWidth > 0 && botao.offsetHeight > 0);
                console.log('Botão habilitado:', !botao.disabled);
                console.log('Has click listener:', !!botao.hasClickListener);
                console.log('Container de itens atual:', document.getElementById('itensContainer')?.children.length || 'não encontrado');
                console.log('ItemCounter atual:', window.itemCounter || 'não definido');
                
                // Testar clique programático
                console.log('🖱️ Testando clique programático...');
                const itensBefore = document.getElementById('itensContainer')?.children.length || 0;
                console.log('Itens antes do clique:', itensBefore);
                
                botao.click();
                
                setTimeout(() => {
                    const itensAfter = document.getElementById('itensContainer')?.children.length || 0;
                    console.log('Itens após clique:', itensAfter);
                    console.log('Novo item adicionado:', itensAfter > itensBefore ? 'SIM' : 'NÃO');
                }, 100);
            }
            return botao;
        }

        // Função de debug para testar Enter
        function testarEnterKeydown() {
            console.log('🧪 TESTE: Simulando tecla Enter');
            const container = document.getElementById('itensContainer');
            const primeiroItem = container?.querySelector('.item-row');
            
            if (primeiroItem) {
                const descricaoInput = primeiroItem.querySelector('input[name="itemDescricao"]');
                if (descricaoInput) {
                    // Preencher campos para teste
                    descricaoInput.value = 'Teste Item';
                    const quantidadeInput = primeiroItem.querySelector('input[name="itemQuantidade"]');
                    const unidadeSelect = primeiroItem.querySelector('select[name="itemUnidade"]');
                    const precoInput = primeiroItem.querySelector('input[name="itemPrecoUnitario"]');
                    
                    if (quantidadeInput) quantidadeInput.value = '1';
                    if (unidadeSelect) unidadeSelect.value = 'UN';
                    if (precoInput) precoInput.value = '10.00';
                    
                    console.log('Campos preenchidos para teste. Simulando Enter...');
                    
                    // Focar no campo e simular Enter
                    descricaoInput.focus();
                    
                    const enterEvent = new KeyboardEvent('keydown', {
                        key: 'Enter',
                        code: 'Enter',
                        bubbles: true,
                        cancelable: true
                    });
                    
                    descricaoInput.dispatchEvent(enterEvent);
                    console.log('Evento Enter disparado');
                }
            } else {
                console.log('❌ Nenhum item encontrado para testar');
            }
        }

        function calcularTotalItem(itemId) {
            const quantidadeElement = document.getElementById(`itemQuantidade_${itemId}`);
            const precoElement = document.getElementById(`itemPrecoUnitario_${itemId}`);
            const totalElement = document.getElementById(`itemTotal_${itemId}`);
            
            if (quantidadeElement && precoElement && totalElement) {
                const quantidade = parseFloat(quantidadeElement.value) || 0;
                const precoUnitario = parseFloat(precoElement.value) || 0;
                const total = quantidade * precoUnitario;
                totalElement.value = `R$ ${total.toFixed(2).replace('.', ',')}`;
                calcularTotalGeral();
            }
        }

        function calcularTotalGeral() {
            let totalGeral = 0;
            
            // Somar com base nos itens armazenados no array
            if (window.itensAdicionados && window.itensAdicionados.length > 0) {
                totalGeral = window.itensAdicionados.reduce((total, item) => total + item.subtotal, 0);
            }
            
            // Atualizar elementos da interface
            const totalGeralElement = document.getElementById('totalGeral');
            const summaryTotalElement = document.getElementById('summary-total');
            
            if (totalGeralElement) {
                totalGeralElement.textContent = `R$ ${totalGeral.toLocaleString('pt-BR', {minimumFractionDigits: 2})}`;
            }
            
            if (summaryTotalElement) {
                summaryTotalElement.textContent = `R$ ${totalGeral.toLocaleString('pt-BR', {minimumFractionDigits: 2})}`;
            }
            
            console.log('💰 Total geral calculado:', totalGeral);
        }
        
        function salvarPedido() {
            console.log('💾 Iniciando processo de salvamento do pedido...');
            
            // ==========================================================
            // VALIDAÇÕES ANTES DO SALVAMENTO
            // ==========================================================
            const fornecedorSelect = document.getElementById('fornecedorPedido');
            const placaInput = document.getElementById('placaVeiculo');
            const veiculoInput = document.getElementById('nomeVeiculo');
            const observacoesInput = document.getElementById('observacoesPedido');
            const perfilUsuario = localStorage.getItem('perfil');
            
            // Verificar se os elementos existem antes de prosseguir
            if (!fornecedorSelect) {
                console.error('❌ Elemento fornecedorPedido não encontrado');
                showToast('Erro interno: Campos do formulário não encontrados.', 'error', 'Erro');
                return;
            }
            
            // Validação 1: Fornecedor selecionado
            if (!fornecedorSelect.value) {
                showToast('Selecione um fornecedor antes de salvar o pedido.', 'error', 'Validação');
                fornecedorSelect.focus();
                return;
            }
            
            // Verificar grupo do fornecedor para aplicar validações específicas
            const selectedOption = fornecedorSelect.options[fornecedorSelect.selectedIndex];
            const estadoFornecedor = selectedOption ? selectedOption.dataset.estado : '';
            const grupoFornecedor = selectedOption ? selectedOption.dataset.grupo : '';
            
            // ==========================================================
            // VALIDAÇÕES ESPECÍFICAS POR GRUPO DE FORNECEDOR
            // ==========================================================
            console.log('🔍 Aplicando validações específicas por grupo');
            console.log('📋 Estado do fornecedor:', estadoFornecedor);
            console.log('📋 Grupo do fornecedor:', grupoFornecedor);
            
            // Validação para fornecedores do grupo OBRIGATÓRIO (qualquer estado)
            if (grupoFornecedor === 'OBRIGATORIO') {
                // Verifica se a placa OU o nome do veículo estão vazios
                if (!placaInput.value.trim() || !veiculoInput.value.trim()) {
                    showToast('Fornecedores do Grupo Obrigatório: Placa e Nome do Veículo são obrigatórios.', 'error', 'Validação Grupo');
                    if (!veiculoInput.value.trim()) veiculoInput.focus();
                    else if (!placaInput.value.trim()) placaInput.focus();
                    return;
                }
                
                // Validação de formato da placa para grupo OBRIGATÓRIO
                if (placaInput && placaInput.value.trim() && placaInput.classList.contains('invalid')) {
                    if (perfilUsuario !== 'admin') {
                        showToast('Fornecedores do Grupo Obrigatório: Placa deve estar em formato válido.', 'error', 'Validação Grupo');
                        placaInput.focus();
                        return;
                    } else {
                        console.warn(`⚠️ Admin salvando pedido (Grupo Obrigatório) com placa inválida: "${placaInput.value}"`);
                    }
                }
            } 
            // Validação para fornecedores do grupo LIVRE (qualquer estado)
            else if (grupoFornecedor === 'LIVRE') {
                // Verifica se as observações estão vazias
                if (!observacoesInput.value.trim()) {
                    showToast('Fornecedores do Grupo Livre: Campo Observações é obrigatório.', 'error', 'Validação Grupo');
                    observacoesInput.focus();
                    return;
                }
                
                // Validação de formato da placa para grupo LIVRE (se a placa foi preenchida)
                if (placaInput && placaInput.value.trim() && placaInput.classList.contains('invalid')) {
                    if (perfilUsuario !== 'admin') {
                        showToast('Fornecedores do Grupo Livre: Se informada, a placa deve estar em formato válido.', 'error', 'Validação Grupo');
                        placaInput.focus();
                        return;
                    } else {
                        console.warn(`⚠️ Admin salvando pedido (Grupo Livre) com placa inválida: "${placaInput.value}"`);
                    }
                }
                
                console.log('📝 Grupo LIVRE: Observações obrigatórias, placa opcional mas com validação de formato');
            }
            // Para fornecedores sem grupo definido
            else {
                // Validação de formato da placa para fornecedores sem grupo (se a placa foi preenchida)
                if (placaInput && placaInput.value.trim() && placaInput.classList.contains('invalid')) {
                    if (perfilUsuario !== 'admin') {
                        showToast('Se informada, a placa deve estar em formato válido.', 'error', 'Validação');
                        placaInput.focus();
                        return;
                    } else {
                        console.warn(`⚠️ Admin salvando pedido (sem grupo) com placa inválida: "${placaInput.value}"`);
                    }
                }
                
                console.log('📝 Sem grupo: Apenas validação de formato de placa se informada');
            }
            
            console.log('✅ Validações específicas por grupo aprovadas:', grupoFornecedor);
            // ==========================================================
            // FIM DAS VALIDAÇÕES ESPECÍFICAS POR GRUPO
            // ==========================================================
            
            // Validação de campos obrigatórios - CONDICIONAL baseada apenas no GRUPO do fornecedor
            const camposObrigatoriosBase = [
                { id: 'fornecedorPedido', nome: 'Fornecedor' }
            ];
            
            // Usar a variável grupoFornecedor já declarada acima
            console.log('📋 Grupo do fornecedor (validação campos):', grupoFornecedor);
            
            // Adicionar campos obrigatórios baseados APENAS no grupo
            if (grupoFornecedor === 'OBRIGATORIO') {
                // Para grupo OBRIGATÓRIO: placa e veículo obrigatórios (observações NÃO obrigatórias)
                camposObrigatoriosBase.push(
                    { id: 'placaVeiculo', nome: 'Placa do Veículo' },
                    { id: 'nomeVeiculo', nome: 'Nome do Veículo' }
                );
                console.log('📋 Grupo OBRIGATÓRIO: Placa e veículo obrigatórios');
            } else if (grupoFornecedor === 'LIVRE') {
                // Para grupo LIVRE: apenas observações obrigatórias
                camposObrigatoriosBase.push(
                    { id: 'observacoesPedido', nome: 'Observações' }
                );
                console.log('📋 Grupo LIVRE: Apenas observações obrigatórias');
            } else {
                // Para fornecedores sem grupo definido: apenas fornecedor obrigatório
                console.log('📋 Sem grupo definido: Apenas fornecedor obrigatório');
            }
            
            let camposFaltando = [];
            camposObrigatoriosBase.forEach(campo => {
                const elemento = document.getElementById(campo.id);
                if (elemento && (!elemento.value || elemento.value.trim() === '')) {
                    camposFaltando.push(campo.nome);
                }
            });
            
            if (camposFaltando.length > 0) {
                showGlobalModal('Erro', `Por favor, preencha os seguintes campos obrigatórios: ${camposFaltando.join(', ')}`);
                return;
            }

            // Coletar itens usando o array de itens adicionados
            const itens = [];
            if (window.itensAdicionados && window.itensAdicionados.length > 0) {
                window.itensAdicionados.forEach(item => {
                    itens.push({
                        descricao: item.descricao,
                        quantidade: item.quantidade,
                        unidade: item.unidade,
                        precoUnitario: item.precoUnitario,
                        totalItem: item.subtotal
                    });
                });
            }

            if (itens.length === 0) {
                showGlobalModal('Erro', 'Adicione pelo menos um item ao pedido.');
                return;
            }

            // Desabilitar botão para evitar duplo clique
            const salvarButton = document.getElementById('salvarPedidoButton');
            if (salvarButton) {
                salvarButton.disabled = true;
                salvarButton.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Gerando número...';
            }
            
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            const idDaEmpresa = empresaAtual.id || empresaAtual.codigo;
            
            // Debug: Verificar dados da empresa no salvamento
            console.log('🔍 Debug SavePedido - Empresa atual:', empresaAtual);
            console.log('🔍 Debug SavePedido - ID da empresa:', idDaEmpresa);
            
            if (!idDaEmpresa) {
                showToast('Erro: Não foi possível identificar a empresa. Faça login novamente.', 'error');
                if (salvarButton) {
                    salvarButton.disabled = false;
                    salvarButton.innerHTML = '<i class="fas fa-check mr-2"></i>Finalizar e Salvar Pedido';
                }
                return;
            }
            
            // Primeiro, gerar o número do pedido
            google.script.run
                .withSuccessHandler(response => {
                    console.log('📋 Resposta da geração de número:', response);
                    
                    // Verificar se o número foi gerado corretamente
                    let numeroPedido;
                    if (typeof response === 'string' || typeof response === 'number') {
                        numeroPedido = response.toString();
                    } else if (response && (response.numero || response.numeroPedido)) {
                        numeroPedido = (response.numero || response.numeroPedido).toString();
                    } else {
                        console.error('❌ Formato de resposta inesperado para número do pedido:', response);
                        showToast('Erro ao gerar número do pedido. Formato de resposta inválido.', 'error', 'Erro de Numeração');
                        if (salvarButton) {
                            salvarButton.disabled = false;
                            salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Pedido';
                        }
                        return;
                    }
                    
                    console.log('📋 Número do pedido processado:', numeroPedido);
                    
                    // Verificar se o número é válido
                    if (!numeroPedido || numeroPedido === 'undefined' || numeroPedido === 'null') {
                        console.error('❌ Número do pedido inválido:', numeroPedido);
                        showToast('Erro: Número do pedido inválido gerado pelo sistema.', 'error', 'Erro de Numeração');
                        if (salvarButton) {
                            salvarButton.disabled = false;
                            salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Pedido';
                        }
                        return;
                    }
                    
                    // Atualizar o campo visual
                    const numeroPedidoElement = document.getElementById('numeroPedido');
                    if (numeroPedidoElement) {
                        numeroPedidoElement.value = numeroPedido;
                    }
                    
                    // Atualizar status do botão
                    if (salvarButton) {
                        salvarButton.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Salvando...';
                    }
                    
                    // VERIFICAÇÃO FINAL antes de montar os dados
                    console.log('🔍 VERIFICAÇÃO FINAL - numeroPedido antes de montar dados:', numeroPedido);
                    console.log('🔍 VERIFICAÇÃO FINAL - tipo da variável:', typeof numeroPedido);
                    
                    // Calcular o total geral dos itens
                    const totalGeral = itens.reduce((total, item) => total + item.totalItem, 0);
                    console.log('💰 Total geral calculado para salvamento:', totalGeral);
                    
                    // Montar dados do pedido com o número gerado
                    const dataParaSalvar = document.getElementById('dataPedido')?.value || '';
                    console.log('📅 Data que será enviada para salvamento:', dataParaSalvar);
                    console.log('📅 Hora atual local:', obterDataHoraLocal());
                    
                    // Converter data do campo para data/hora completa local
                    let dataCompleta = dataParaSalvar;
                    if (dataParaSalvar && !dataParaSalvar.includes(':')) {
                        // Se é só data (YYYY-MM-DD), adicionar hora atual
                        const agora = new Date();
                        const hora = String(agora.getHours()).padStart(2, '0');
                        const minuto = String(agora.getMinutes()).padStart(2, '0');
                        const segundo = String(agora.getSeconds()).padStart(2, '0');
                        dataCompleta = `${dataParaSalvar} ${hora}:${minuto}:${segundo}`;
                        console.log('📅 Data convertida para hora completa:', dataCompleta);
                    }
                    
                    const dadosPedido = {
                        numeroPedido: String(numeroPedido), // Garantir que é string
                        data: dataCompleta,
                        fornecedor: document.getElementById('fornecedorPedido')?.value || '',
                        nomeVeiculo: document.getElementById('nomeVeiculo')?.value || '',
                        placaVeiculo: document.getElementById('placaVeiculo')?.value || '',
                        observacoes: document.getElementById('observacoesPedido')?.value || '',
                        empresaId: idDaEmpresa,
                        totalGeral: totalGeral, // Incluir o total geral calculado
                        itens: itens
                    };
                    
                    // Forçar atribuição se por algum motivo não funcionou
                    if (!dadosPedido.numeroPedido || dadosPedido.numeroPedido === 'undefined') {
                        dadosPedido.numeroPedido = String(numeroPedido);
                        console.warn('⚠️ Forçando atribuição do numeroPedido:', dadosPedido.numeroPedido);
                    }

                    console.log('📋 Dados do pedido montados:', dadosPedido);
                    console.log('📋 Número do pedido para salvamento:', numeroPedido);
                    console.log('🔍 VERIFICAÇÃO CRÍTICA - dadosPedido.numeroPedido:', dadosPedido.numeroPedido);
                    console.log('🔍 ENVIANDO PARA BACKEND - Objeto completo:', JSON.stringify(dadosPedido, null, 2));
                    
                    // Verificação final antes do envio
                    if (!dadosPedido.numeroPedido || dadosPedido.numeroPedido === 'undefined') {
                        console.error('❌ ERRO CRÍTICO: numeroPedido está undefined no dadosPedido!');
                        showToast('Erro crítico: Número do pedido perdido. Tentando novamente...', 'error');
                        if (salvarButton) {
                            salvarButton.disabled = false;
                            salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Pedido';
                        }
                        return;
                    }

                    // Agora salvar o pedido
                    google.script.run
                        .withSuccessHandler(response => {
                            console.log('📋 Resposta do salvamento:', response);
                            
                            // Verificar diferentes formatos de resposta
                            if (response && (response.status === 'success' || response.status === 'ok')) {
                                // Backend retornou sucesso, mas pode ter o bug do "undefined"
                                const mensagem = response.message || '';
                                
                                if (mensagem.includes('undefined')) {
                                    console.warn('⚠️ Backend retornou "undefined" na mensagem, mas salvou com sucesso');
                                    showToast(`Pedido ${numeroPedido} salvo com sucesso! (corrigindo mensagem do backend)`, 'success', 'Pedido Criado');
                                } else {
                                    // Tentar obter o número do pedido de diferentes formas
                                    const numeroFinal = response.numeroPedido || response.numero || numeroPedido;
                                    showToast(`Pedido ${numeroFinal} salvo com sucesso!`, 'success', 'Pedido Criado');
                                }
                                
                                setTimeout(() => {
                                    loadPageContent('Menu', setupMenuScreen);
                                }, 2000);
                            } else if (typeof response === 'string' && response.includes('sucesso')) {
                                // Se a resposta for uma string de sucesso
                                if (response.includes('undefined')) {
                                    showToast(`Pedido ${numeroPedido} salvo com sucesso! (corrigindo backend)`, 'success', 'Pedido Criado');
                                } else {
                                    showToast(`Pedido ${numeroPedido} salvo com sucesso!`, 'success', 'Pedido Criado');
                                }
                                setTimeout(() => {
                                    loadPageContent('Menu', setupMenuScreen);
                                }, 2000);
                            } else {
                                // Erro no salvamento
                                const mensagemErro = response?.message || response || 'Erro desconhecido no salvamento';
                                showToast('Erro ao salvar: ' + mensagemErro, 'error', 'Falha no Salvamento');
                                console.error('❌ Erro detalhado no salvamento:', response);
                            }
                            
                            if (salvarButton) {
                                salvarButton.disabled = false;
                                salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Pedido';
                            }
                        })
                        .withFailureHandler(err => {
                            showToast('Erro de comunicação: ' + err.message, 'error', 'Falha na Conexão');
                            if (salvarButton) {
                                salvarButton.disabled = false;
                                salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Pedido';
                            }
                        })
                        .salvarPedido(dadosPedido);
                })
                .withFailureHandler(err => {
                    console.error('Erro ao gerar número do pedido:', err);
                    showToast('Erro ao gerar número do pedido: ' + err.message, 'error', 'Erro de Numeração');
                    if (salvarButton) {
                        salvarButton.disabled = false;
                        salvarButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Pedido';
                    }
                })
                .getProximoNumeroPedido(idDaEmpresa);
        }

        // ===============================================
        // FUNÇÃO PARA SALVAR RASCUNHO
        // ===============================================
        function salvarRascunho() {
            console.log('📝 Iniciando salvamento de rascunho...');
            
            // Validações mínimas para rascunho (apenas fornecedor e pelo menos 1 item)
            const fornecedorSelect = document.getElementById('fornecedorPedido');
            
            if (!fornecedorSelect.value) {
                showToast('Selecione um fornecedor para salvar o rascunho.', 'warning', 'Validação');
                fornecedorSelect.focus();
                return;
            }

            // Verificar se há pelo menos um item adicionado
            if (!window.itensAdicionados || window.itensAdicionados.length === 0) {
                showToast('Adicione pelo menos um item com descrição para salvar o rascunho.', 'warning', 'Validação');
                return;
            }

            // Desabilitar botão para evitar duplo clique
            const rascunhoButton = document.getElementById('salvarRascunhoButton');
            if (rascunhoButton) {
                rascunhoButton.disabled = true;
                rascunhoButton.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Salvando...';
            }

            // Coletar dados do rascunho usando o array de itens
            const itens = window.itensAdicionados.map(item => ({
                descricao: item.descricao,
                quantidade: item.quantidade,
                unidade: item.unidade,
                precoUnitario: item.precoUnitario,
                totalItem: item.subtotal
            }));

            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            
            // Debug: Verificar dados da empresa
            console.log('🔍 Debug - Empresa atual:', empresaAtual);
            console.log('🔍 Debug - ID da empresa:', empresaAtual.id);
            console.log('🔍 Debug - Código da empresa:', empresaAtual.codigo);
            
            // Usar id ou codigo como fallback
            const empresaId = empresaAtual.id || empresaAtual.codigo;
            if (!empresaId) {
                showToast('Erro: Não foi possível identificar a empresa. Faça login novamente.', 'error');
                return;
            }

            console.log('🔍 Debug - ID da empresa usado:', empresaId);
            
            // Calcular total geral dos itens
            const totalGeral = window.itensAdicionados ? 
                window.itensAdicionados.reduce((total, item) => total + item.subtotal, 0) : 0;
            
            console.log('💰 Debug - Total geral calculado:', totalGeral);
            
            // Verificar se estamos editando um rascunho existente
            const editandoRascunho = localStorage.getItem('editandoRascunho') === 'true';
            const rascunhoId = localStorage.getItem('rascunhoId');
            
            console.log('🔄 Modo edição:', editandoRascunho);
            console.log('🔄 ID do rascunho:', rascunhoId);
            
            // Dados do rascunho
            const dadosRascunho = {
                data: document.getElementById('dataPedido')?.value || '',
                fornecedor: fornecedorSelect.value,
                nomeVeiculo: document.getElementById('nomeVeiculo')?.value || '',
                placaVeiculo: document.getElementById('placaVeiculo')?.value || '',
                observacoes: document.getElementById('observacoesPedido')?.value || '',
                empresa: empresaId,
                itens: itens,
                totalGeral : totalGeral,
                status: 'RASCUNHO',
                dataUltimaEdicao: obterDataHoraLocal()
            };
            
            // Se estiver editando, incluir o ID do rascunho
            if (editandoRascunho && rascunhoId) {
                dadosRascunho.rascunhoId = rascunhoId;  // Usar 'rascunhoId' em vez de 'id'
                console.log('📝 Atualizando rascunho existente:', rascunhoId);
            } else {
                console.log('📝 Criando novo rascunho');
            }

            // Debug: Mostrar dados que serão enviados
            console.log('📤 Debug - Dados do rascunho a serem enviados:', dadosRascunho);
            console.log('📤 Debug - Quantidade de itens:', itens.length);
            console.log('📤 Debug - Empresa ID final:', empresaId);

            // Salvar rascunho no backend
            google.script.run
                .withSuccessHandler(response => {
                    console.log('✅ Resposta do backend:', response);
                    if (response.status === 'success') {
                        const isEdicao = editandoRascunho && rascunhoId;
                        const mensagem = isEdicao ? 
                            `Rascunho atualizado com sucesso! ID: ${rascunhoId}` :
                            `Rascunho salvo com sucesso! ID: ${response.rascunhoId}`;
                        
                        showToast(mensagem, 'success', 'Rascunho Salvo');
                        
                        // Mostrar indicador de rascunho na tela
                        const draftStatus = document.getElementById('draft-status');
                        if (draftStatus) {
                            draftStatus.classList.remove('hidden');
                            const idFinal = isEdicao ? rascunhoId : response.rascunhoId;
                            draftStatus.innerHTML = `<i class="fas fa-edit mr-2"></i>Rascunho - ID: ${idFinal}`;
                        }
                        
                        // Salvar ID do rascunho para futuras edições
                        const idParaSalvar = isEdicao ? rascunhoId : response.rascunhoId;
                        localStorage.setItem('rascunhoAtual', idParaSalvar);
                        
                        // Marcar que agora está em modo de edição
                        if (!isEdicao) {
                            localStorage.setItem('editandoRascunho', 'true');
                            localStorage.setItem('rascunhoId', response.rascunhoId);
                        }
                        
                    } else {
                        showToast('Erro ao salvar rascunho: ' + response.message, 'error', 'Falha no Salvamento');
                    }
                    
                    // Reabilitar botão
                    if (rascunhoButton) {
                        rascunhoButton.disabled = false;
                        rascunhoButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Rascunho';
                    }
                })
                .withFailureHandler(err => {
                    console.error('❌ Erro detalhado ao salvar rascunho:', err);
                    console.error('❌ Tipo do erro:', typeof err);
                    console.error('❌ Erro stringify:', JSON.stringify(err));
                    
                    // Tentar extrair mais informações do erro
                    const errorMessage = err.message || err.toString() || 'Erro desconhecido';
                    showToast('Erro ao salvar rascunho: ' + errorMessage, 'error', 'Falha no Salvamento');
                    
                    if (rascunhoButton) {
                        rascunhoButton.disabled = false;
                        rascunhoButton.innerHTML = '<i class="fas fa-save mr-2"></i>Salvar Rascunho';
                    }
                })
                .salvarRascunho(dadosRascunho);
        }

        // ===============================================
        // LÓGICA DA PÁGINA DE CADASTRO DE FORNECEDOR
        // ===============================================
        function setupCadastroFornecedorScreen() {
            // Guarda de proteção (garanta que o ID corresponde ao seu container principal)
            if (!document.getElementById('form-fornecedor')) {
                console.warn('Formulário de fornecedor não encontrado');
                return;
            }
            
            console.log("Inicializando a página de Cadastro de Fornecedor...");

            // --- CONFIGURAÇÃO DOS BOTÕES PRINCIPAIS ---
            const saveButton = document.getElementById('saveButton');
            const cancelButton = document.getElementById('cancelButton');
            const consultarCnpjBtn = document.getElementById('consultarCnpjBtn');

            if (saveButton) {
                saveButton.addEventListener('click', salvarFornecedor);
            }
            
            if (cancelButton) {
                cancelButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            if (consultarCnpjBtn) {
                consultarCnpjBtn.addEventListener('click', handleConsultarCnpj);
            }

            // Carregar dados iniciais
            preencherCombosFornecedor();
            carregarEstados();

            // --- LÓGICA PARA FORÇAR MAIÚSCULAS ---
            // Aplica a regra para todos os inputs de texto e textareas
            document.querySelectorAll('#form-fornecedor input[type="text"], #form-fornecedor textarea').forEach(input => {
                // Adiciona o "ouvinte" para o evento 'blur' (quando o usuário sai do campo)
                input.addEventListener('blur', forceUppercase);
            });

            // --- LÓGICA UNIFICADA PARA O CAMPO CNPJ ---
            const cnpjInput = document.getElementById('cnpj');

            if (cnpjInput) {
                // Listener 1: Máscara automática enquanto o usuário digita
                cnpjInput.addEventListener('input', (e) => {
                    e.target.value = formatarCpfCnpj(e.target.value);
                });

                // Listener 2: Validação visual quando o usuário sai do campo
                cnpjInput.addEventListener('blur', (e) => {
                    const cnpjDigitado = e.target.value;
                    if (cnpjDigitado) {
                        const isValid = validarCnpj(cnpjDigitado);
                        e.target.classList.toggle('valid', isValid);
                        e.target.classList.toggle('invalid', !isValid);
                    } else {
                        // Limpa as classes se o campo estiver vazio
                        e.target.classList.remove('valid', 'invalid');
                    }
                });
            }
        }

          // Função para carregar estados dinamicamente do backend
        function carregarEstados() {
            const estadoSelect = document.getElementById('estado');
            if (!estadoSelect) return;

            // Mostrar loading
            estadoSelect.innerHTML = '<option value="">Carregando estados...</option>';

            google.script.run
                .withSuccessHandler(estados => {
                    console.log('[carregarEstados] Estados recebidos:', estados);
                    populateSelect('estado', estados, 'Selecione o Estado');
                })
                .withFailureHandler(error => {
                    console.error('[carregarEstados] Erro ao buscar estados:', error);
                    showToast('Erro ao carregar estados.', 'error', 'Erro');
                  populateSelect('estado', estadosLocal, 'Selecione o Estado');
                })
                .getEstados();
        }


        function setupBuscaScreen() {
            console.log('🔧 Inicializando nova tela de busca de pedidos...');
            
            // Elementos da nova interface
            const searchButton = document.getElementById('searchButton');
            const mainSearch = document.getElementById('mainSearch');
            const backButton = document.getElementById('backToMenuFromSearchButton');
            const toggleAdvancedBtn = document.getElementById('toggleAdvancedFilters');
            const advancedFiltersDiv = document.getElementById('advancedFilters');
            const clearFiltersButton = document.getElementById('clearFiltersButton');
            const dateStart = document.getElementById('dateStart');
            const dateEnd = document.getElementById('dateEnd');
            const plateSearch = document.getElementById('plateSearch');
            const userSearch = document.getElementById('userSearch');
            const resultsBody = document.getElementById('searchResultsBody');

            // Event listener para busca principal
            if (searchButton) {
                searchButton.addEventListener('click', realizarBuscaAvancada);
            }

            if (mainSearch) {
                mainSearch.addEventListener('keypress', (e) => {
                    if (e.key === 'Enter') {
                        realizarBuscaAvancada();
                    }
                });
            }

            // Toggle filtros avançados
            if (toggleAdvancedBtn && advancedFiltersDiv) {
                toggleAdvancedBtn.addEventListener('click', () => {
                    const icon = toggleAdvancedBtn.querySelector('i');
                    advancedFiltersDiv.classList.toggle('hidden');
                    if (icon) {
                        icon.classList.toggle('fa-chevron-down');
                        icon.classList.toggle('fa-chevron-up');
                    }
                });
            }

            // Limpar filtros
            if (clearFiltersButton) {
                clearFiltersButton.addEventListener('click', () => {
                    if (mainSearch) mainSearch.value = '';
                    if (dateStart) dateStart.value = '';
                    if (dateEnd) dateEnd.value = '';
                    if (plateSearch) plateSearch.value = '';
                    if (userSearch) userSearch.value = '';
                    
                    // Mostrar mensagem inicial em vez de buscar
                    if (resultsBody) {
                        resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-gray-500 py-12"><div class="flex flex-col items-center"><i class="fas fa-search fa-3x mb-4 text-gray-300"></i><p class="text-lg font-medium">Use os filtros acima para buscar pedidos</p><p class="text-sm text-gray-400 mt-2">Digite um número de pedido, nome do fornecedor ou use os filtros avançados</p></div></td></tr>';
                    }
                    
                    showToast('Filtros limpos', 'success');
                });
            }

            // Voltar ao menu
            if (backButton) {
                backButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            // Configurar modal de impressão
            setupPrintModal();
            
            // Mostrar mensagem inicial em vez de carregar todos os pedidos
            if (resultsBody) {
                resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-gray-500 py-12"><div class="flex flex-col items-center"><i class="fas fa-search fa-3x mb-4 text-gray-300"></i><p class="text-lg font-medium">Use os filtros acima para buscar pedidos</p><p class="text-sm text-gray-400 mt-2">Digite um número de pedido, nome do fornecedor ou use os filtros avançados</p></div></td></tr>';
            }
        }

        function setupDashboardScreen() {
            const backButton = document.getElementById('backToMenuFromDashboard');
            
            if (backButton) {
                backButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            carregarDadosDashboard();
        }

        function setupRelatoriosScreen() {
            const backButton = document.getElementById('backToMenuFromReportsButton');
            const generatePdfButton = document.getElementById('generatePdfButton');
            const generateXlsButton = document.getElementById('generateXlsButton');

            if (backButton) {
                backButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            if (generatePdfButton) {
                generatePdfButton.addEventListener('click', () => gerarRelatorio('pdf'));
            }

            if (generateXlsButton) {
                generateXlsButton.addEventListener('click', () => gerarRelatorio('xlsx'));
            }

            carregarFornecedoresRelatorio();
        }

        function setupGerenciarRascunhosScreen() {
            console.log('🔧 Inicializando tela de gerenciamento de rascunhos...');
            
            const backButton = document.getElementById('backToMenuFromRascunhosButton');
            const messageElement = document.getElementById('rascunhosMessage');
            const tableBody = document.getElementById('rascunhosTableBody');
            const emptyState = document.getElementById('emptyStateRascunhos');
            const totalRascunhos = document.getElementById('totalRascunhos');
            
            let allRascunhos = [];
            
            // Event listener para voltar ao menu
            if (backButton) {
                backButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }
            
            // Função para mostrar mensagem
            function showMessage(text, type = 'success') {
                if (messageElement) {
                    messageElement.textContent = text;
                    messageElement.className = `text-center text-sm font-medium min-h-[20px] mt-2 transition-opacity duration-300 ${type === 'success' ? 'text-green-600' : 'text-red-600'}`;
                    setTimeout(() => { messageElement.textContent = ''; }, 3000);
                }
            }
            
            // Função para renderizar tabela de rascunhos
            function renderRascunhosTable(rascunhos) {
                if (!tableBody) return;
                
                console.log('🔍 Debug - Rascunhos recebidos:', rascunhos);
                console.log('🔍 Debug - Primeiro rascunho completo:', rascunhos[0]);
                
                tableBody.innerHTML = '';
                
                if (rascunhos.length === 0) {
                    if (emptyState) emptyState.classList.remove('hidden');
                    if (totalRascunhos) totalRascunhos.textContent = 'Nenhum rascunho encontrado';
                    return;
                }
                
                if (emptyState) emptyState.classList.add('hidden');
                if (totalRascunhos) totalRascunhos.textContent = `${rascunhos.length} rascunho(s) encontrado(s)`;
                
                rascunhos.forEach(rascunho => {
                    console.log(`🔍 Debug - Processando rascunho ${rascunho.id}:`);
                    console.log(`  - dataUltimaEdicao:`, rascunho.dataUltimaEdicao);
                    console.log(`  - tipo dataUltimaEdicao:`, typeof rascunho.dataUltimaEdicao);
                    
                    const dataFormatada = new Date(rascunho.data).toLocaleDateString('pt-BR');
                    
                    let ultimaEdicao = 'N/A';
                    if (rascunho.dataUltimaEdicao && rascunho.dataUltimaEdicao.trim() !== '') {
                        try {
                            const dataEdicao = new Date(rascunho.dataUltimaEdicao);
                            if (!isNaN(dataEdicao.getTime())) {
                                ultimaEdicao = dataEdicao.toLocaleString('pt-BR');
                            }
                        } catch (e) {
                            console.warn(`⚠️ Erro ao formatar data de última edição para rascunho ${rascunho.id}:`, e);
                        }
                    }
                        
                    console.log(`  - ultimaEdicao formatada:`, ultimaEdicao);
                    
                    const row = document.createElement('tr');
                    row.className = 'hover:bg-gray-50';
                    row.innerHTML = `
                        <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${rascunho.id}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${dataFormatada}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${rascunho.fornecedor}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${rascunho.estadoFornecedor || 'N/A'}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${rascunho.itens?.length || 0} item(s)</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${ultimaEdicao}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-center text-sm font-medium space-x-2">
                            <button onclick="editarRascunho('${rascunho.id}')" 
                                    class="text-blue-600 hover:text-blue-900" title="Editar">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button onclick="finalizarRascunho('${rascunho.id}')" 
                                    class="text-green-600 hover:text-green-900" title="Finalizar como Pedido">
                                <i class="fas fa-check-circle"></i>
                            </button>
                            <button onclick="excluirRascunho('${rascunho.id}')" 
                                    class="text-red-600 hover:text-red-900" title="Excluir">
                                <i class="fas fa-trash"></i>
                            </button>
                        </td>
                    `;
                    tableBody.appendChild(row);
                });
            }
            
            // Função para carregar rascunhos do backend
            function carregarRascunhos() {
                showMessage('Carregando rascunhos...', 'success');
                
                // Primeiro, vamos testar a comunicação
                console.log('🔧 Testando comunicação com o backend...');
                google.script.run
                    .withSuccessHandler(testResponse => {
                        console.log('✅ Teste de comunicação bem-sucedido:', testResponse);
                        
                        // Agora tentar buscar rascunhos
                        const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
                        console.log('🔍 Empresa para busca:', empresaAtual);
                        
                        google.script.run
                            .withSuccessHandler(response => {
                                console.log('📋 Resposta do backend para buscarRascunhos:', response);
                                
                                // Verificar se response não é null/undefined
                                if (!response) {
                                    console.error('❌ Resposta null/undefined do backend');
                                    showMessage('Erro: Nenhuma resposta do servidor. Verifique a função backend.', 'error');
                                    renderRascunhosTable([]);
                                    return;
                                }
                                
                                // Verificar se response tem a estrutura esperada
                                if (typeof response !== 'object') {
                                    console.error('❌ Resposta inválida do backend:', typeof response, response);
                                    showMessage('Erro: Resposta inválida do servidor.', 'error');
                                    renderRascunhosTable([]);
                                    return;
                                }
                                
                                if (response.status === 'success') {
                                    allRascunhos = response.rascunhos || [];
                                    renderRascunhosTable(allRascunhos);
                                    carregarFornecedoresFiltro();
                                    showMessage(`${allRascunhos.length} rascunho(s) carregado(s)`, 'success');
                                } else if (response.status === 'error') {
                                    const errorMsg = response.message || 'Erro desconhecido no servidor';
                                    showMessage('Erro ao carregar rascunhos: ' + errorMsg, 'error');
                                    renderRascunhosTable([]);
                                } else {
                                    console.error('❌ Status desconhecido na resposta:', response.status);
                                    showMessage('Erro: Status desconhecido na resposta do servidor.', 'error');
                                    renderRascunhosTable([]);
                                }
                            })
                            .withFailureHandler(error => {
                                console.error('❌ Erro ao carregar rascunhos:', error);
                                showMessage('Erro de comunicação ao buscar rascunhos: ' + (error.message || error), 'error');
                                renderRascunhosTable([]);
                            })
                            .buscarRascunhosCorrigida(empresaAtual.id || empresaAtual.codigo);
                    })
                    .withFailureHandler(error => {
                        console.error('❌ Falha no teste de comunicação:', error);
                        showMessage('Erro: Falha na comunicação com o servidor. Verifique a configuração.', 'error');
                        renderRascunhosTable([]);
                    })
                    .testarComunicacao();
            }
            
            // Função para carregar fornecedores no filtro
            function carregarFornecedoresFiltro() {
                const filtroFornecedor = document.getElementById('filtroFornecedor');
                if (!filtroFornecedor) return;
                
                // Limpar e adicionar opção padrão
                filtroFornecedor.innerHTML = '<option value="">Todos os fornecedores</option>';
                
                // Primeiro, adicionar fornecedores únicos dos rascunhos
                const fornecedoresRascunhos = [...new Set(allRascunhos.map(r => r.fornecedor).filter(f => f))];
                
                // Carregar todos os fornecedores da base de dados
                const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
                google.script.run
                    .withSuccessHandler(response => {
                        let todosFornecedores = [];
                        
                        // getFornecedoresList retorna array direto
                        if (Array.isArray(response)) {
                            todosFornecedores = response;
                        }
                        
                        // Combinar fornecedores dos rascunhos com todos os fornecedores
                        const fornecedoresCombinados = new Set([
                            ...fornecedoresRascunhos,
                            ...todosFornecedores.map(f => f.razao || f.text || f.razaoSocial || f.nome)
                        ]);
                        
                        // Adicionar ao select
                        Array.from(fornecedoresCombinados)
                            .filter(f => f && f.trim())
                            .sort()
                            .forEach(fornecedor => {
                                const option = document.createElement('option');
                                option.value = fornecedor;
                                option.textContent = fornecedor;
                                filtroFornecedor.appendChild(option);
                            });
                    })
                    .withFailureHandler(error => {
                        console.error('Erro ao carregar fornecedores:', error);
                        // Em caso de erro, usar apenas fornecedores dos rascunhos
                        fornecedoresRascunhos.forEach(fornecedor => {
                            const option = document.createElement('option');
                            option.value = fornecedor;
                            option.textContent = fornecedor;
                            filtroFornecedor.appendChild(option);
                        });
                    })
                    .getFornecedoresList();
            }
            
            // Event listeners para filtros
            ['filtroFornecedor', 'filtroDataInicio', 'filtroDataFim'].forEach(filtroId => {
                const filtro = document.getElementById(filtroId);
                if (filtro) {
                    filtro.addEventListener('change', aplicarFiltros);
                }
            });
            
            
            // Função para aplicar filtros
            function aplicarFiltros() {
                const filtroFornecedor = document.getElementById('filtroFornecedor')?.value || '';
                const filtroDataInicio = document.getElementById('filtroDataInicio')?.value || '';
                const filtroDataFim = document.getElementById('filtroDataFim')?.value || '';
                
                let rascunhosFiltrados = allRascunhos;
                
                if (filtroFornecedor) {
                    rascunhosFiltrados = rascunhosFiltrados.filter(r => r.fornecedor === filtroFornecedor);
                }
                
                if (filtroDataInicio) {
                    rascunhosFiltrados = rascunhosFiltrados.filter(r => r.data >= filtroDataInicio);
                }
                
                if (filtroDataFim) {
                    rascunhosFiltrados = rascunhosFiltrados.filter(r => r.data <= filtroDataFim);
                }
                
                renderRascunhosTable(rascunhosFiltrados);
            }
            
            // Carregar rascunhos na inicialização
            carregarRascunhos();
        }

        function setupGerenciarFornecedoresScreen() {
            const backButton = document.getElementById('backToMenuFromSuppliersButton');
            const addButton = document.getElementById('addNewFornecedorBtn');

            if (backButton) {
                backButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            if (addButton) {
                addButton.addEventListener('click', () => {
                    loadPageContent('CadastroFornecedor', setupCadastroFornecedorScreen);
                });
            }

            carregarListaFornecedores();
        }

        function setupTrocarSenhaScreen() {
            const btnSalvar = document.getElementById('btnSalvarSenha');
            const btnVoltar = document.getElementById('btnVoltarMenuSenha');

            if (btnSalvar) {
                btnSalvar.addEventListener('click', trocarSenha);
            }

            if (btnVoltar) {
                btnVoltar.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }
        }

        function setupPedidoSalvoScreen(numeroPedido) {
            const voltarButton = document.getElementById('voltarMenuPrincipalButton');
            const printButton = document.getElementById('printPedidoButton');

            if (voltarButton) {
                voltarButton.addEventListener('click', () => {
                    loadPageContent('Menu', setupMenuScreen);
                });
            }

            if (printButton) {
                printButton.addEventListener('click', () => {
                    abrirImpressaoPedido(numeroPedido);
                });
            }
        }

        function abrirImpressaoPedido(numeroPedido) {
            if (!numeroPedido) {
                showGlobalModal('Erro', 'Número do pedido não fornecido para impressão.');
                return;
            }
            
            showGlobalModal('Aguarde', `Preparando impressão do pedido ${numeroPedido}...`, 'info');

            google.script.run
                .withSuccessHandler(pedidoData => {
                    if (!pedidoData) {
                        showGlobalModal('Erro', 'Dados do pedido não encontrados para impressão.');
                        return;
                    }
                    
                    document.getElementById('global-modal').classList.add('hidden');

                    const pageHtml = getPageTemplate('PedidoImpressao');
                    const printWindow = window.open('', '_blank');
                    
                    if (!printWindow) {
                        showGlobalModal("Pop-up bloqueado!", "Por favor, permita pop-ups para este site para poder imprimir o pedido.");
                        return;
                    }

                    printWindow.document.write(pageHtml);
                    printWindow.document.close();

                    printWindow.onload = function() {
                        preencherDadosImpressao(printWindow.document, pedidoData);
                        
                        setTimeout(() => {
                            printWindow.focus();
                            printWindow.print();
                        }, 500);
                    };
                })
                .withFailureHandler(err => {
                    showGlobalModal('Erro', `Erro ao carregar dados do pedido: ${err.message}`);
                })
                .getDadosPedidoParaImpressao(numeroPedido);
        }

        function preencherDadosImpressao(doc, pedidoData) {
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            
            // Dados da empresa
            doc.getElementById('empresaNomePrint').textContent = empresaAtual.nome || '';
            doc.getElementById('empresaEnderecoPrint').textContent = pedidoData.enderecoEmpresa || '';
            doc.getElementById('empresaCidadeUfPrint').textContent = pedidoData.cidadeUfEmpresa || '';
            doc.getElementById('empresaCnpjPrint').textContent = pedidoData.cnpjEmpresa || '';
            doc.getElementById('empresaEmailPrint').textContent = `Email: ${pedidoData.emailEmpresa || ''}`;
            doc.getElementById('empresaTelefonePrint').textContent = `Telefone: ${pedidoData.telefoneEmpresa || ''}`;

            // Dados do pedido
            doc.getElementById('numeroPedidoPrint').textContent = pedidoData.numeroDoPedido || '';
            
            const dataPedidoObj = new Date(pedidoData.data + 'T12:00:00');
            const formattedDate = dataPedidoObj.toLocaleDateString('pt-BR');
            const now = new Date();
            const formattedTime = now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
            doc.getElementById('dataEmissaoPrint').textContent = `${formattedDate} ${formattedTime}`;

            // Dados do fornecedor
            doc.getElementById('fornecedorPrint').textContent = pedidoData.fornecedor || '';
            doc.getElementById('enderecoFornecedorPrint').textContent = pedidoData.enderecoFornecedor || '';
            doc.getElementById('formaPagamentoPrint').textContent = pedidoData.formaPagamentoFornecedor || '';
            doc.getElementById('cnpjFornecedorPrint').textContent = pedidoData.cnpjFornecedor || '';
            doc.getElementById('condicaoPagamentoPrint').textContent = pedidoData.condicaoPagamentoFornecedor || '';
            
            // Dados do veículo
            doc.getElementById('placaVeiculoPrint').textContent = pedidoData.placaVeiculo || 'N/A';
            doc.getElementById('nomeVeiculoPrint').textContent = pedidoData.nomeVeiculo || 'N/A';
            
            // Total e observações
            doc.getElementById('totalGeralPrint').textContent = `R$ ${parseFloat(pedidoData.totalGeral || 0).toFixed(2).replace('.', ',')}`;
            doc.getElementById('observacoesPedidoPrint').textContent = pedidoData.observacoes || 'Sem observações.';

            // Itens
            const printItemsTableBody = doc.getElementById('printItemsTableBody');
            printItemsTableBody.innerHTML = '';
            if (pedidoData.itens && Array.isArray(pedidoData.itens)) {
                pedidoData.itens.forEach((item, index) => {
                    const row = printItemsTableBody.insertRow();
                    row.insertCell(0).textContent = index + 1;
                    row.insertCell(1).textContent = item.descricao;
                    row.insertCell(2).textContent = item.unidade || '';
                    row.insertCell(3).textContent = item.quantidade;
                    row.insertCell(4).textContent = `R$ ${parseFloat(item.precoUnitario || 0).toFixed(2).replace('.', ',')}`;
                    row.insertCell(5).textContent = `R$ ${parseFloat(item.totalItem || 0).toFixed(2).replace('.', ',')}`;
                });
            }
            
            // Dados do usuário
            const nomeUsuario = localStorage.getItem('nome') || 'Usuário Desconhecido';
            const perfilUsuario = localStorage.getItem('perfil') || 'Sem Função';
            doc.getElementById('usuarioLogadoPrint').textContent = nomeUsuario.toUpperCase();
            doc.getElementById('funcaoUsuarioLogadoPrint').textContent = perfilUsuario.toUpperCase();
        }



        function realizarBuscaAvancada() {
            console.log('🔍 Realizando busca avançada...');
            
            const mainSearch = document.getElementById('mainSearch');
            const dateStart = document.getElementById('dateStart');
            const dateEnd = document.getElementById('dateEnd');
            const plateSearch = document.getElementById('plateSearch');
            const userSearch = document.getElementById('userSearch');
            const resultsBody = document.getElementById('searchResultsBody');
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');

            // Coletar filtros
            const filtros = {
                query: mainSearch ? mainSearch.value.trim() : '',
                dataInicial: dateStart ? dateStart.value : '',
                dataFinal: dateEnd ? dateEnd.value : '',
                placa: plateSearch ? plateSearch.value.trim() : '',
                usuario: userSearch ? userSearch.value.trim() : ''
            };

            console.log('🔍 Filtros aplicados:', filtros);

            // Verificar se pelo menos um filtro foi preenchido
            const temFiltros = filtros.query || filtros.dataInicial || filtros.dataFinal || filtros.placa || filtros.usuario;
            
            if (!temFiltros) {
                if (resultsBody) {
                    resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-gray-500 py-12"><div class="flex flex-col items-center"><i class="fas fa-search fa-3x mb-4 text-gray-300"></i><p class="text-lg font-medium">Use os filtros acima para buscar pedidos</p><p class="text-sm text-gray-400 mt-2">Digite um número de pedido, nome do fornecedor ou use os filtros avançados</p></div></td></tr>';
                }
                showToast('Digite pelo menos um critério de busca', 'info');
                return;
            }

            // Limpar tabela e mostrar loading
            if (resultsBody) {
                resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-gray-500 py-8"><div class="spinner mx-auto mb-2"></div>Buscando pedidos...</td></tr>';
            }

            console.log('🔍 Iniciando busca no backend...');
            console.log('📝 Parâmetros da busca:');
            console.log('   - Query:', filtros.query || '');
            console.log('   - Empresa ID:', empresaAtual.id);
            
            google.script.run
                .withSuccessHandler(response => {
                    console.log('✅ Resposta da busca recebida:', response);
                    if (response.status === 'success' && response.data && response.data.length > 0) {
                        // Filtrar os resultados no frontend com base nos filtros avançados
                        const resultadosFiltrados = filtrarPedidosAvancado(response.data, filtros);
                        
                        if (resultadosFiltrados.length > 0) {
                            preencherTabelaBusca(resultadosFiltrados, resultsBody);
                            showToast(`${resultadosFiltrados.length} pedido(s) encontrado(s)`, 'success');
                        } else {
                            if (resultsBody) {
                                resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-gray-500 py-8">Nenhum pedido encontrado com os filtros aplicados</td></tr>';
                            }
                            showToast('Nenhum pedido encontrado com os filtros aplicados', 'info');
                        }
                    } else {
                        console.log('⚠️ Busca retornou resultado vazio ou com erro');
                        console.log('   - Status:', response.status);
                        console.log('   - Data:', response.data);
                        console.log('   - Message:', response.message);
                        
                        if (resultsBody) {
                            resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-gray-500 py-8">Nenhum pedido encontrado</td></tr>';
                        }
                        showToast('Nenhum pedido encontrado', 'info');
                    }
                })
                .withFailureHandler(err => {
                    console.error('🔍 Erro na busca:', err);
                    if (resultsBody) {
                        resultsBody.innerHTML = '<tr><td colspan="6" class="text-center text-red-500 py-8">Erro ao buscar pedidos. Tente novamente.</td></tr>';
                    }
                    showToast('Erro na busca: ' + err.message, 'error');
                })
                .buscarPedidos(filtros.query || '', empresaAtual.id);
        }

        function filtrarPedidosAvancado(pedidos, filtros) {
            console.log('🔍 Filtrando pedidos com filtros avançados:', filtros);
            console.log('📊 Total de pedidos recebidos:', pedidos.length);
            
            if (!pedidos || pedidos.length === 0) {
                return [];
            }

            const pedidosFiltrados = pedidos.filter(pedido => {
                // Filtro por placa
                if (filtros.placa && filtros.placa.trim()) {
                    const placaPedido = (pedido.placa || '').toLowerCase();
                    const placaFiltro = filtros.placa.toLowerCase().trim();
                    if (!placaPedido.includes(placaFiltro)) {
                        return false;
                    }
                }

                // Filtro por usuário
                if (filtros.usuario && filtros.usuario.trim()) {
                    const usuarioPedido = (pedido.usuario || pedido.usuarioResponsavel || '').toLowerCase();
                    const usuarioFiltro = filtros.usuario.toLowerCase().trim();
                    if (!usuarioPedido.includes(usuarioFiltro)) {
                        return false;
                    }
                }

                // Filtro por data inicial
                if (filtros.dataInicial && filtros.dataInicial.trim()) {
                    try {
                        const dataInicial = new Date(filtros.dataInicial + 'T00:00:00');
                        const dataPedido = normalizarDataPedido(pedido.data);
                        if (dataPedido && dataPedido < dataInicial) {
                            return false;
                        }
                    } catch (e) {
                        console.warn('Erro ao comparar data inicial:', e);
                    }
                }

                // Filtro por data final
                if (filtros.dataFinal && filtros.dataFinal.trim()) {
                    try {
                        const dataFinal = new Date(filtros.dataFinal + 'T23:59:59');
                        const dataPedido = normalizarDataPedido(pedido.data);
                        if (dataPedido && dataPedido > dataFinal) {
                            return false;
                        }
                    } catch (e) {
                        console.warn('Erro ao comparar data final:', e);
                    }
                }

                return true;
            });

            console.log('📊 Pedidos após filtros avançados:', pedidosFiltrados.length);
            return pedidosFiltrados;
        }

        function normalizarDataPedido(dataInput) {
            console.log('📅 Normalizando data do pedido:', dataInput);
            
            if (!dataInput) {
                console.warn('📅 Data do pedido está vazia ou nula');
                return null;
            }

            // Se já é um objeto Date válido
            if (dataInput instanceof Date && !isNaN(dataInput)) {
                return dataInput;
            }

            // Se é uma string, tentar vários formatos
            if (typeof dataInput === 'string') {
                let dataString = dataInput.trim();
                
                // Formato "YYYY-MM-DD HH:mm:ss" - precisa de conversão especial
                if (dataString.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
                    // IMPORTANTE: Não adicionar Z para manter horário local brasileiro
                    // Converter para formato ISO sem indicar UTC
                    const dataISO = dataString.replace(' ', 'T');
                    const data = new Date(dataISO);
                    if (!isNaN(data)) {
                        console.log('📅 Data convertida do formato "YYYY-MM-DD HH:mm:ss":', data.toLocaleString('pt-BR'));
                        return data;
                    }
                }
                
                // Formato ISO padrão (YYYY-MM-DD ou YYYY-MM-DDTHH:mm:ss)
                if (dataString.match(/^\d{4}-\d{2}-\d{2}T/)) {
                    const data = new Date(dataString);
                    if (!isNaN(data)) {
                        console.log('📅 Data convertida do formato ISO:', data.toLocaleString('pt-BR'));
                        return data;
                    }
                }
                
                // Formato de data simples (YYYY-MM-DD)
                if (dataString.match(/^\d{4}-\d{2}-\d{2}$/)) {
                    const data = new Date(dataString + 'T12:00:00.000Z');
                    if (!isNaN(data)) {
                        console.log('📅 Data convertida do formato YYYY-MM-DD:', data.toLocaleString('pt-BR'));
                        return data;
                    }
                }
                
                // Formato brasileiro (DD/MM/YYYY)
                if (dataString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
                    const [dia, mes, ano] = dataString.split('/');
                    const data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia), 12, 0, 0);
                    if (!isNaN(data)) {
                        console.log('📅 Data convertida do formato brasileiro:', data.toLocaleString('pt-BR'));
                        return data;
                    }
                }
                
                // Tentar parseamento direto apenas como último recurso
                const dataTentativa = new Date(dataString);
                if (!isNaN(dataTentativa)) {
                    console.log('📅 Data convertida por parseamento direto:', dataTentativa.toLocaleString('pt-BR'));
                    return dataTentativa;
                }
            }

            console.warn('📅 Não foi possível normalizar a data:', dataInput);
            return null;
        }

        function preencherTabelaBusca(pedidos, tbody) {
            console.log('📊 Preenchendo tabela com', pedidos.length, 'pedidos');
            if (!tbody) return;
            
            tbody.innerHTML = '';
            
            pedidos.forEach(pedido => {
                console.log('📊 Processando pedido:', pedido);
                
                const row = document.createElement('tr');
                row.className = 'hover:bg-gray-50';
                
                // Verificar se o pedido está cancelado
                const isCanceled = pedido.status === 'Cancelado';
                if (isCanceled) {
                    row.classList.add('canceled-row');
                }

                // Verificar se pode editar (menos de 1 hora)
                const podeEditar = verificarSePermiteEdicao(pedido);
                
                const statusClass = isCanceled ? 'bg-red-200 text-red-800' : 
                                   pedido.status === 'Ativo' ? 'bg-green-100 text-green-800' : 
                                   'bg-gray-100 text-gray-800';
                
                const valorFormatado = `R$ ${parseFloat(pedido.totalGeral || 0).toFixed(2).replace('.', ',')}`;
                
                // Melhor tratamento para formatação de data
                let dataFormatada = 'N/A';
                try {
                    const dataNormalizada = normalizarDataPedido(pedido.data);
                    if (dataNormalizada && !isNaN(dataNormalizada.getTime())) {
                        dataFormatada = dataNormalizada.toLocaleDateString('pt-BR');
                    }
                } catch (e) {
                    console.warn('Erro ao formatar data do pedido:', pedido.data, e);
                }

                // Normalizar campo do número do pedido - aceitar vários formatos possíveis do backend
                const numeroPedido = pedido.numeroDoPedido || pedido.numeroPedido || pedido.numero || 'N/A';

                // Verificar se pedido tem dados válidos (evitar mostrar pedidos vazios)
                if (!numeroPedido || numeroPedido === 'N/A' || !pedido.fornecedor) {
                    console.warn('Pedido com dados incompletos ignorado:', pedido);
                    return; // Pular este pedido
                }

                row.innerHTML = `
                    <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${numeroPedido}</td>
                    <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${dataFormatada}</td>
                    <td class="px-6 py-4 text-sm text-gray-500">${pedido.fornecedor || 'N/A'}</td>
                    <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900 text-right font-bold">${valorFormatado}</td>
                    <td class="px-6 py-4 whitespace-nowrap text-center">
                        <span class="px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${statusClass}">
                            ${pedido.status || 'Em Aberto'}
                        </span>
                    </td>
                    <td class="px-6 py-4 whitespace-nowrap text-center text-sm font-medium space-x-2">
                        <button onclick="abrirImpressaoPedido('${numeroPedido}')" title="Imprimir" class="text-blue-600 hover:text-blue-900">
                            <i class="fas fa-print"></i>
                        </button>
                        <button onclick="editarPedido('${numeroPedido}')" title="Editar" 
                                class="text-yellow-600 hover:text-yellow-900 ${!podeEditar || isCanceled ? 'opacity-50 cursor-not-allowed' : ''}" 
                                ${!podeEditar || isCanceled ? 'disabled' : ''}>
                            <i class="fas fa-pencil-alt"></i>
                        </button>
                        <button onclick="cancelarPedido('${numeroPedido}')" title="Cancelar" 
                                class="text-red-600 hover:text-red-900 ${isCanceled ? 'opacity-50 cursor-not-allowed' : ''}" 
                                ${isCanceled ? 'disabled' : ''}>
                            <i class="fas fa-times-circle"></i>
                        </button>
                    </td>
                `;
                
                tbody.appendChild(row);
            });
        }

        function setupPrintModal() {
            const printModal = document.getElementById('print-modal');
            const closePrintModalBtn = document.getElementById('close-print-modal');

            if (closePrintModalBtn) {
                closePrintModalBtn.addEventListener('click', () => {
                    if (printModal) {
                        printModal.classList.add('hidden');
                        printModal.classList.remove('flex');
                    }
                });
            }

            // Fechar modal clicando fora
            if (printModal) {
                printModal.addEventListener('click', (e) => {
                    if (e.target === printModal) {
                        printModal.classList.add('hidden');
                        printModal.classList.remove('flex');
                    }
                });
            }
        }

        function abrirModalImpressao(numeroPedido) {
            console.log('🖨️ Abrindo modal de impressão para pedido:', numeroPedido);
            
            const printModal = document.getElementById('print-modal');
            const printContent = document.getElementById('print-content');
            const printPedidoId = document.getElementById('print-pedido-id');

            if (printPedidoId) {
                printPedidoId.textContent = numeroPedido;
            }

            // Verificar se o pedido está cancelado para adicionar marca d'água
            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'success' && response.data && response.data.length > 0) {
                        const pedido = response.data[0];
                        if (printContent) {
                            if (pedido.status === 'Cancelado') {
                                printContent.classList.add('print-watermark');
                            } else {
                                printContent.classList.remove('print-watermark');
                            }
                        }
                    }
                })
                .withFailureHandler(err => {
                    console.error('Erro ao verificar status do pedido:', err);
                })
                .buscarPedidos(numeroPedido, JSON.parse(localStorage.getItem('empresaSelecionada') || '{}').id);

            if (printModal) {
                printModal.classList.remove('hidden');
                printModal.classList.add('flex');
            }
        }

        function cancelarPedido(numeroPedido) {
            if (confirm(`Tem a certeza que deseja cancelar o pedido #${numeroPedido}?`)) {
                console.log('❌ Cancelando pedido:', numeroPedido);
                
                showToast('Cancelando pedido...', 'info');
                
                // Por enquanto, apenas simular o cancelamento atualizando localmente
                // Em uma implementação real, você faria a chamada para o backend
                setTimeout(() => {
                    showToast(`Pedido #${numeroPedido} cancelado com sucesso`, 'success');
                    realizarBuscaAvancada(); // Recarregar resultados
                }, 1000);
            }
        }

          // Função para verificar se um pedido pode ser editado (1 hora após criação)
        function verificarSePermiteEdicao(pedido) {
            try {
                console.log('🔍 Verificando permissão de edição para pedido:', pedido);
                console.log('📋 Campos de data disponíveis:', {
                    dataHoraCriacao: pedido.dataHoraCriacao,
                    dataUltimaEdicao: pedido.dataUltimaEdicao,
                    data: pedido.data,
                    dataCriacao: pedido.dataCriacao,
                    timestamp: pedido.timestamp
                });
                
                // Log detalhado de TODOS os campos do pedido para identificar onde está a data/hora real
                console.log('🔍 TODOS os campos do pedido para análise:');
                for (const [campo, valor] of Object.entries(pedido)) {
                    if (typeof valor === 'string' && (valor.includes('2025') || valor.includes('11:59') || valor.includes(':'))) {
                        console.log(`📍 Campo "${campo}":`, valor);
                    }
                }
                
                const agora = new Date();
                let dataCriacao;
                
                // PRIORIZAR o campo correto: 'Data Hora Criacao' (com espaços)
                if (pedido['Data Hora Criacao']) {
                    let dataString = pedido['Data Hora Criacao'];
                    console.log('🎯 CAMPO CORRETO encontrado - "Data Hora Criacao":', dataString);
                    
                    // Se está no formato "YYYY-MM-DD HH:mm:ss", converter para ISO sem UTC
                    if (dataString.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
                        dataString = dataString.replace(' ', 'T'); // SEM adicionar .000Z
                        console.log('🔧 Convertendo Data Hora Criacao (horário local):', pedido['Data Hora Criacao'], '→', dataString);
                    }
                    
                    dataCriacao = new Date(dataString);
                    console.log('📅 Data de criação processada (Data Hora Criacao):', dataCriacao.toLocaleString('pt-BR'));
                }
                // Tentar usar dataHoraCriacao primeiro
                else if (pedido.dataHoraCriacao) {
                    // Corrigir formato de data para preservar horário local brasileiro
                    let dataString = pedido.dataHoraCriacao;
                    
                    // Se está no formato "YYYY-MM-DD HH:mm:ss", converter para ISO sem UTC
                    if (dataString.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
                        dataString = dataString.replace(' ', 'T'); // SEM adicionar .000Z
                        console.log('🔧 Convertendo formato de data (horário local):', pedido.dataHoraCriacao, '→', dataString);
                    }
                    
                    dataCriacao = new Date(dataString);
                    console.log('📅 Data de criação processada:', dataCriacao.toLocaleString('pt-BR'));
                }
                // Tentar usar "Data Criacao" (com espaço)
                else if (pedido['Data Criacao']) {
                    let dataString = pedido['Data Criacao'];
                    console.log('🔧 Encontrado campo "Data Criacao":', dataString);
                    
                    // Se está no formato "YYYY-MM-DD HH:mm:ss", converter para ISO sem UTC
                    if (dataString.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
                        dataString = dataString.replace(' ', 'T'); // SEM adicionar .000Z
                        console.log('🔧 Convertendo Data Criacao (horário local):', pedido['Data Criacao'], '→', dataString);
                    }
                    
                    dataCriacao = new Date(dataString);
                    console.log('📅 Data de criação processada (Data Criacao):', dataCriacao.toLocaleString('pt-BR'));
                }
                // Tentar usar dataCriacao (sem espaço)
                else if (pedido.dataCriacao) {
                    let dataString = pedido.dataCriacao;
                    console.log('🔧 Encontrado campo "dataCriacao":', dataString);
                    
                    // Se está no formato "YYYY-MM-DD HH:mm:ss", converter para ISO sem UTC
                    if (dataString.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
                        dataString = dataString.replace(' ', 'T'); // SEM adicionar .000Z
                        console.log('🔧 Convertendo dataCriacao (horário local):', pedido.dataCriacao, '→', dataString);
                    }
                    
                    dataCriacao = new Date(dataString);
                    console.log('📅 Data de criação processada (dataCriacao):', dataCriacao.toLocaleString('pt-BR'));
                }
                // Tentar usar o campo "data" se contém data/hora completa
            try {
                console.log('🔍 Verificando permissão de edição para pedido:', pedido);
                // Usar sempre o campo 'data' (coluna C) para a data de criação
                const agora = new Date();
                let dataCriacao = null;
                if (pedido.data) {
                    dataCriacao = normalizarDataPedido(pedido.data);
                }
                if (!dataCriacao || isNaN(dataCriacao.getTime())) {
                    console.warn('⚠️ Data de criação inválida para pedido:', pedido);
                    return false;
                }
                // Calcular diferença em minutos
                const diffMs = agora.getTime() - dataCriacao.getTime();
                const diffMin = diffMs / (1000 * 60);
                // Só libera edição se for até 60 minutos
                const podeEditar = diffMin >= 0 && diffMin <= 60;
                console.log(`⏱️ Pedido criado há ${diffMin.toFixed(2)} minutos. Pode editar?`, podeEditar);
                return podeEditar;
            } catch (e) {
                console.error('Erro ao verificar permissão de edição:', e);
                return false;
            }
        }
                    return false;
                }
                return verificarSePermiteEdicao(this.dados);
            }
        };

        // Função para editar pedido - MELHORADA
        function editarPedido(numeroPedido) {
            console.log('✏️ Iniciando edição do pedido:', numeroPedido);
            const empresaAtual = JSON.parse(localStorage.getItem('empresaSelecionada') || '{}');
            
            showToast('Carregando dados do pedido para edição...', 'info');
            
            google.script.run
                .withSuccessHandler(response => {
                    console.log('📋 Resposta da busca para edição:', response);
                    if (response.status === 'success' && response.data && response.data.length > 0) {
                        const pedidoData = response.data[0];
                        
                        // Verificar se ainda pode editar
                        if (!verificarSePermiteEdicao(pedidoData)) {
                            showToast('Este pedido não pode mais ser editado. Edição disponível apenas por 1 hora após a criação.', 'warning');
                            return;
                        }
                        
                        // Armazenar dados no sistema de edição
                        EdicaoPedido.setDados(pedidoData);
                        
                        // Reutilizar a tela de pedido existente em modo edição
                        loadPageContent('Pedido', () => setupPedidoScreen(true, pedidoData));
                        showToast('Pedido carregado para edição. Você tem 1 hora para salvar as alterações.', 'success');
                        
                    } else {
                        showToast('Erro ao carregar dados do pedido.', 'error');
                    }
                })
                .withFailureHandler(err => {
                    showToast('Erro de comunicação ao buscar pedido.', 'error');
                    console.error('Erro ao buscar pedido para edição:', err);
                })
                .buscarPedidos(numeroPedido, empresaAtual.id);
        }

        function trocarSenha() {
            const senhaAtual = document.getElementById('senhaAtual').value.trim();
            const novaSenha = document.getElementById('novaSenha').value.trim();
            const confirmaSenha = document.getElementById('confirmaSenha').value.trim();
            const usuarioLogado = localStorage.getItem('usuarioLogado');

            if (!senhaAtual || !novaSenha || !confirmaSenha) {
                showToast('Preencha todos os campos', 'error', 'msgTrocaSenha');
                return;
            }

            if (novaSenha !== confirmaSenha) {
                showToast('As senhas não coincidem', 'error', 'msgTrocaSenha');
                return;
            }

            if (novaSenha.length < 6) {
                showToast('A nova senha deve ter pelo menos 6 caracteres', 'error', 'msgTrocaSenha');
                return;
            }

            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'ok') {
                        showToast('Senha alterada com sucesso!', 'success', 'Segurança');
                        setTimeout(() => {
                            loadPageContent('Menu', setupMenuScreen);
                        }, 2000);
                    } else {
                        showToast(response.message, 'error', 'Erro na Alteração');
                    }
                })
                .withFailureHandler(err => {
                    showToast('Erro ao alterar senha: ' + err.message, 'error', 'Falha na Conexão');
                })
                .alterarSenhaUsuario(usuarioLogado, senhaAtual, novaSenha);
        }

        // ===============================================
        // FUNÇÕES DE FORNECEDOR
        // ===============================================
        
        function preencherCombosFornecedor() {
            showToast('Carregando dados do fornecedor...', 'info', 'Carregando');
            console.log("[preencherCombosFornecedor] Iniciando busca de dados...");

            google.script.run
                .withSuccessHandler(codigo => {
                    const codigoInput = document.getElementById('codigo');
                    if (codigoInput) codigoInput.value = codigo;
                    console.log("[preencherCombosFornecedor] Código recebido:", codigo);
                    showToast('Dados carregados com sucesso!', 'success', 'Sucesso');
                })
                .withFailureHandler(error => {
                    console.error('[preencherCombosFornecedor] Erro ao buscar próximo código:', error);
                    showToast('Erro ao carregar código do fornecedor.', 'error', 'Erro');
                })
                .getProximoCodigoFornecedor();

            google.script.run
                .withSuccessHandler(condicoes => {
                    console.log('[preencherCombosFornecedor] Condições de pagamento recebidas:', condicoes);
                    populateSelect('condicao', condicoes, 'Selecione a Condição');
                })
                .withFailureHandler(error => {
                    console.error('[preencherCombosFornecedor] Erro ao buscar condições de pagamento:', error);
                    showToast('Erro ao carregar condições de pagamento.', 'error', 'Erro');
                })
                .getCondicoesPagamento();

            google.script.run
                .withSuccessHandler(formas => {
                    console.log('[preencherCombosFornecedor] Formas de pagamento recebidas:', formas);
                    populateSelect('forma', formas, 'Selecione a Forma');
                })
                .withFailureHandler(error => {
                    console.error('[preencherCombosFornecedor] Erro ao buscar formas de pagamento:', error);
                    showToast('Erro ao carregar formas de pagamento.', 'error', 'Erro');
                })
                .getFormasPagamento();
        }

        function salvarFornecedor() {
            const form = document.getElementById('form-fornecedor');
            const saveButton = document.getElementById('saveButton');

            if (!form || !saveButton) {
                console.error("Elementos do formulário de fornecedor não encontrados para salvar.");
                showToast("Erro interno: Elementos do formulário ausentes.", "error", "Erro Interno");
                return;
            }

            if (!form.checkValidity()) {
                showToast('Por favor, preencha todos os campos obrigatórios!', 'error', 'Campos Obrigatórios');
                form.reportValidity();
                return;
            }

            const fornecedor = {
                codigo: document.getElementById('codigo').value.trim(),
                razao: document.getElementById('razao').value.trim(),
                fantasia: document.getElementById('fantasia').value.trim(),
                cnpj: document.getElementById('cnpj').value.trim(),
                endereco: document.getElementById('endereco').value.trim(),
                estado: document.getElementById('estado').value.trim(),
                condicao: document.getElementById('condicao').value.trim(),
                forma: document.getElementById('forma').value.trim()
            };

            saveButton.disabled = true;
            saveButton.innerText = 'Salvando...';
            showToast('Salvando fornecedor...', 'info', 'Processando');

            google.script.run
                .withSuccessHandler(res => {
                    saveButton.disabled = false;
                    saveButton.innerText = 'Salvar';

                    if (res.status === 'ok') {
                        showToast(res.message || 'Fornecedor salvo com sucesso!', 'success', 'Sucesso');
                        form.reset();
                        // Buscar novo código após salvar
                        google.script.run
                            .withSuccessHandler(newCodigo => {
                                const codigoInput = document.getElementById('codigo');
                                if (codigoInput) codigoInput.value = newCodigo;
                            })
                            .withFailureHandler(error => {
                                console.error('Erro ao buscar novo código após salvamento:', error);
                            })
                            .getProximoCodigoFornecedor();
                    } else {
                        showToast(res.message || 'Erro ao salvar fornecedor. Tente novamente.', 'error', 'Erro no Salvamento');
                    }
                })
                .withFailureHandler(error => {
                    saveButton.disabled = false;
                    saveButton.innerText = 'Salvar';
                    console.error('Erro na comunicação com o servidor (salvarFornecedor):', error);
                    showToast('Ocorreu um erro inesperado ao salvar. Tente novamente.', 'error', 'Erro de Comunicação');
                })
                .salvarFornecedor(fornecedor);
        }
                
        /**
         * Formata um valor para o padrão de CPF (se tiver até 11 dígitos) 
         * ou CNPJ (se tiver mais de 11 dígitos).
         */
        function formatarCpfCnpj(valor) {
            // 1. Remove tudo que não for número
            const apenasNumeros = valor.replace(/\D/g, '');

            // 2. Verifica o tamanho e aplica a máscara correta
            if (apenasNumeros.length <= 11) {
                // É um CPF
                return apenasNumeros
                    .replace(/(\d{3})(\d)/, '$1.$2')
                    .replace(/(\d{3})(\d)/, '$1.$2')
                    .replace(/(\d{3})(\d{1,2})$/, '$1-$2');
            } else {
                // É um CNPJ
                return apenasNumeros.slice(0, 14) // Limita a 14 dígitos
                    .replace(/^(\d{2})(\d)/, '$1.$2')
                    .replace(/^(\d{2})\.(\d{3})(\d)/, '$1.$2.$3')
                    .replace(/\.(\d{3})(\d)/, '.$1/$2')
                    .replace(/(\d{4})(\d)/, '$1-$2');
            }
        }

        /**
         * Valida um CNPJ (verifica o formato e os dígitos verificadores).
         * @param {string} cnpj - O CNPJ a ser validado.
         * @returns {boolean} - Retorna true se o CNPJ for válido.
         */
        function validarCnpj(cnpj) {
            if (!cnpj) return false;

            // Remove caracteres especiais e espaços
            const apenasNumeros = cnpj.replace(/[^\d]+/g, '');

            // CNPJ deve ter 14 dígitos
            if (apenasNumeros.length !== 14) return false;

            // Elimina CNPJs inválidos conhecidos (todos os números iguais)
            if (/^(\d)\1+$/.test(apenasNumeros)) return false;

            // Validação dos dígitos verificadores
            let tamanho = apenasNumeros.length - 2;
            let numeros = apenasNumeros.substring(0, tamanho);
            let digitos = apenasNumeros.substring(tamanho);
            let soma = 0;
            let pos = tamanho - 7;

            for (let i = tamanho; i >= 1; i--) {
                soma += numeros.charAt(tamanho - i) * pos--;
                if (pos < 2) pos = 9;
            }
            
            let resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(0)) return false;

            tamanho = tamanho + 1;
            numeros = apenasNumeros.substring(0, tamanho);
            soma = 0;
            pos = tamanho - 7;
            for (let i = tamanho; i >= 1; i--) {
                soma += numeros.charAt(tamanho - i) * pos--;
                if (pos < 2) pos = 9;
            }

            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(1)) return false;

            return true;
        }

        // Função para lidar com o clique do botão "Consultar"
        function handleConsultarCnpj() {
            const cnpjInput = document.getElementById('cnpj');
            const consultarBtn = document.getElementById('consultarCnpjBtn');
            const cnpj = cnpjInput.value;

            if (!validarCnpj(cnpj)) {
                showToast('O CNPJ digitado é inválido.', 'error', 'CNPJ Inválido');
                return;
            }
            
            showToast('Consultando CNPJ...', 'info', 'Consultando');
            consultarBtn.disabled = true;
            consultarBtn.textContent = '...';

            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'ok') {
                        document.getElementById('razao').value = response.data.razaoSocial || '';
                        document.getElementById('fantasia').value = response.data.nomeFantasia || '';
                        document.getElementById('endereco').value = response.data.endereco || '';
                        // Preencher o estado se disponível na resposta da API
                        const estadoSelect = document.getElementById('estado');
                        if (estadoSelect && response.data.uf) {
                            estadoSelect.value = response.data.uf;
                        }
                        showToast('Dados preenchidos com sucesso!', 'success', 'CNPJ Consultado');
                    } else {
                        showToast(response.message, 'error', 'Erro na Consulta');
                    }
                    consultarBtn.disabled = false;
                    consultarBtn.textContent = 'Consultar';
                })
                .withFailureHandler(err => {
                    showToast('Erro de comunicação ao consultar CNPJ.', 'error', 'Erro de Comunicação');
                    consultarBtn.disabled = false;
                    consultarBtn.textContent = 'Consultar';
                })
                .consultarCnpj(cnpj);
        }

        // Função para forçar maiúsculas
        function forceUppercase(event) {
            const input = event.target;
            input.value = input.value.toUpperCase();
        }

        // Função para popular selects
        function populateSelect(selectId, options, placeholder) {
            const select = document.getElementById(selectId);
            if (!select) {
                console.warn(`Select com ID '${selectId}' não encontrado`);
                return;
            }

            select.innerHTML = '';
            
            // Adicionar opção placeholder
            if (placeholder) {
                const defaultOption = document.createElement('option');
                defaultOption.value = '';
                defaultOption.textContent = placeholder;
                defaultOption.disabled = true;
                defaultOption.selected = true;
                select.appendChild(defaultOption);
            }

            // Adicionar opções
            if (Array.isArray(options)) {
                options.forEach(option => {
                    const opt = document.createElement('option');
                    if (typeof option === 'string') {
                        opt.value = option;
                        opt.textContent = option;
                    } else if (option.value !== undefined && option.text !== undefined) {
                        opt.value = option.value;
                        opt.textContent = option.text;
                    } else {
                        console.warn('Formato de opção inválido:', option);
                    }
                    select.appendChild(opt);
                });
            }
        }

        // Funções antigas de fornecedor (manter compatibilidade)
        function carregarCondicoesPagamento() {
            google.script.run
                .withSuccessHandler(condicoes => {
                    populateSelect('condicao', condicoes, 'Selecione a Condição');
                })
                .withFailureHandler(error => {
                    console.error('Erro ao buscar condições de pagamento:', error);
                    showToast('Erro ao carregar condições de pagamento.', 'error', 'Erro');
                })
                .getCondicoesPagamento();
        }

        function carregarFormasPagamento() {
            google.script.run
                .withSuccessHandler(formas => {
                    populateSelect('forma', formas, 'Selecione a Forma');
                })
                .withFailureHandler(error => {
                    console.error('Erro ao buscar formas de pagamento:', error);
                    showToast('Erro ao carregar formas de pagamento.', 'error', 'Erro');
                })
                .getFormasPagamento();
        }

        function consultarCNPJ() {
            handleConsultarCnpj();
        }

        function gerarRelatorio(tipo) {
            showToast(`Gerando relatório ${tipo.toUpperCase()}...`, 'info', 'reportMessage');
        }

        function carregarFornecedoresRelatorio() {
            // Implementar carregamento para relatórios
        }

        function carregarListaFornecedores() {
            showToast('Carregando fornecedores...', 'info', 'fornecedorListMessage');
        }

        // ===============================================
        // FUNÇÕES DE RETRY PARA CARREGAMENTO DOS SELECTS
        // ===============================================

        // Função melhorada para aguardar carregamento dos selects
        async function aguardarSelectsCarregados(maxTentativas = 10, intervalo = 500) {
            console.log("🔍 Aguardando carregamento dos selects...");
            
            for (let tentativa = 1; tentativa <= maxTentativas; tentativa++) {
                const selectFornecedor = document.getElementById('fornecedorPedido');
                const selectVeiculo = document.getElementById('nomeVeiculo');
                
                console.log(`⏳ Tentativa ${tentativa}/${maxTentativas}`);
                
                if (selectFornecedor && selectVeiculo) {
                    const fornecedorCarregado = selectFornecedor.options.length > 1;
                    const veiculoCarregado = selectVeiculo.options.length > 1;
                    
                    console.log(`📊 Fornecedor: ${fornecedorCarregado ? '✅' : '❌'} (${selectFornecedor.options.length} options)`);
                    console.log(`📊 Veículo: ${veiculoCarregado ? '✅' : '❌'} (${selectVeiculo.options.length} options)`);
                    
                    if (fornecedorCarregado && veiculoCarregado) {
                        console.log("✅ Ambos selects carregados com sucesso!");
                        return true;
                    }
                } else {
                    console.log("❌ Elementos select não encontrados no DOM");
                }
                
                // Aguardar antes da próxima tentativa
                await new Promise(resolve => setTimeout(resolve, intervalo));
            }
            
            console.log("⚠️ Timeout: Selects não carregaram no tempo esperado");
            return false;
        }

        // Função de preenchimento com retry
        async function preencherFormularioComRetry(dados) {
            console.log("=== INICIANDO PREENCHIMENTO COM RETRY ===");
            
            try {
                // Aguardar carregamento dos selects
                const selectsCarregados = await aguardarSelectsCarregados();
                
                if (!selectsCarregados) {
                    console.log("🔄 Selects não carregaram, tentando recarregar dados...");
                    // Recarregar fornecedores e veículos
                    await recarregarDadosSelects();
                    
                    // Tentar novamente
                    const segundaTentativa = await aguardarSelectsCarregados(5, 1000);
                    if (!segundaTentativa) {
                        throw new Error("Falha ao carregar selects após retry");
                    }
                }
                
                // Agora preencher com segurança
                await preencherCamposFormulario(dados);
                
            } catch (error) {
                console.error("❌ Erro no preenchimento com retry:", error);
                showGlobalModal("Erro", "Erro ao carregar dados do pedido: " + error.message);
            }
        }

        // Função para recarregar dados dos selects
        async function recarregarDadosSelects() {
            console.log("🔄 Recarregando dados dos selects...");
            
            return Promise.all([
                new Promise((resolve) => {
                    google.script.run
                        .withSuccessHandler(fornecedores => {
                            const selectFornecedor = document.getElementById('fornecedorPedido');
                            if (selectFornecedor) {
                                selectFornecedor.innerHTML = '<option value="">Selecione um Fornecedor</option>';
                                fornecedores.forEach(f => {
                                    const option = document.createElement('option');
                                    option.value = f.razao || f.codigo;
                                    option.textContent = f.razao || f.codigo;
                                    selectFornecedor.appendChild(option);
                                });
                            }
                            resolve(fornecedores);
                        })
                        .withFailureHandler(() => resolve([]))
                        .getFornecedoresList();
                }),
                new Promise((resolve) => {
                    google.script.run
                        .withSuccessHandler(veiculos => {
                            const selectVeiculo = document.getElementById('nomeVeiculo');
                            if (selectVeiculo) {
                                selectVeiculo.innerHTML = '<option value="">Selecione um veículo</option>';
                                console.log('📋 Recarregar - Dados brutos dos veículos:', veiculos);
                                
                                veiculos.forEach((v, index) => {
                                    // Tentar diferentes propriedades para encontrar o nome
                                    let nomeVeiculo = null;
                                    
                                    if (typeof v === 'string' && v.trim() !== '' && v !== 'undefined') {
                                        nomeVeiculo = v.trim();
                                    } else if (typeof v === 'object' && v !== null) {
                                        nomeVeiculo = v.nomeVeiculo || v.nome || v.veiculo || v.descricao;
                                    }
                                    
                                    if (nomeVeiculo && nomeVeiculo.trim() !== '' && nomeVeiculo !== 'undefined') {
                                        const option = document.createElement('option');
                                        option.value = nomeVeiculo.trim();
                                        option.textContent = nomeVeiculo.trim();
                                        selectVeiculo.appendChild(option);
                                    } else {
                                        console.warn(`⚠️ Recarregar - Veículo inválido no índice ${index}:`, v);
                                    }
                                });
                            }
                            resolve(veiculos);
                        })
                        .withFailureHandler(() => resolve([]))
                        .getVeiculosList();
                })
            ]);
        }

        // Função melhorada para preencher campos do formulário
        async function preencherCamposFormulario(dados) {
            console.log("📝 Preenchendo campos do formulário...");
            
            // Preencher campos básicos
            const numeroPedido = document.getElementById('numeroPedido');
            if (numeroPedido) {
                numeroPedido.value = dados.numeroDoPedido || '';
                console.log(`✅ Número do pedido: ${dados.numeroDoPedido}`);
            }

            const dataPedido = document.getElementById('dataPedido');
            if (dataPedido) {
                dataPedido.value = dados.data || '';
                console.log(`✅ Data: ${dados.data}`);
            }

            const placaVeiculo = document.getElementById('placaVeiculo');
            if (placaVeiculo) {
                placaVeiculo.value = dados.placaVeiculo || '';
                console.log(`✅ Placa: ${dados.placaVeiculo}`);
            }

            const observacoes = document.getElementById('observacoesPedido');
            if (observacoes) {
                observacoes.value = dados.observacoes || '';
                console.log(`✅ Observações: ${dados.observacoes || 'Vazio'}`);
            }

            // Preencher total
            const totalGeral = document.getElementById('totalGeral');
            if (totalGeral && dados.totalGeral) {
                const valorFormatado = `R$ ${parseFloat(dados.totalGeral).toFixed(2).replace('.', ',')}`;
                totalGeral.value = valorFormatado;
                console.log(`✅ Total: ${valorFormatado}`);
            }

            // Preencher selects
            const selectFornecedor = document.getElementById('fornecedorPedido');
            if (selectFornecedor && dados.fornecedor) {
                selectFornecedor.value = dados.fornecedor;
                console.log(`✅ Fornecedor selecionado: ${dados.fornecedor}`);
            }

            const selectVeiculo = document.getElementById('nomeVeiculo');
            if (selectVeiculo && dados.nomeVeiculo) {
                selectVeiculo.value = dados.nomeVeiculo;
                console.log(`✅ Veículo selecionado: ${dados.nomeVeiculo}`);
            }

            // Preencher itens
            if (dados.itens && Array.isArray(dados.itens)) {
                console.log(`Preenchendo ${dados.itens.length} itens...`);
                preencherItensContainer(dados.itens);
            }

            // Ativar modo visualização
            desabilitarEdicaoPedido();
            console.log("✅ Preenchimento concluído com sucesso!");
        }

        // Função auxiliar global para impressão direta (usada nos cards do dashboard)
        function imprimirPedidoDireto(numeroPedido, empresaId) {
            abrirImpressaoPedido(numeroPedido);
        }
        
        // ===============================================
        // FUNÇÕES GLOBAIS PARA GERENCIAR RASCUNHOS
        // ===============================================
        
        function editarRascunho(rascunhoId) {
            console.log('✏️ Editando rascunho:', rascunhoId);
            
            showToast('Carregando rascunho para edição...', 'info');
            
            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'success' && response.rascunho) {
                        // Salvar dados do rascunho para edição
                        localStorage.setItem('editandoRascunho', 'true');
                        localStorage.setItem('rascunhoId', rascunhoId);
                        localStorage.setItem('dadosRascunho', JSON.stringify(response.rascunho));
                        
                        // Ir para tela de pedido
                        loadPageContent('Pedido', setupPedidoScreen);
                        
                        // Aguardar carregamento da tela e preencher dados
                        setTimeout(() => {
                            preencherFormularioComRascunho(response.rascunho);
                            mostrarIndicadorRascunho(rascunhoId);
                        }, 500);
                        
                    } else {
                        showToast('Erro ao carregar rascunho: ' + response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    console.error('Erro ao buscar rascunho:', error);
                    showToast('Erro de comunicação: ' + error.message, 'error');
                })
                .buscarRascunhoPorId(rascunhoId);
        }
        
        function finalizarRascunho(rascunhoId) {
            console.log('✅ Finalizando rascunho:', rascunhoId);
            
            const modal = document.getElementById('global-modal');
            if (!modal) {
                console.error('Modal global não encontrado');
                return;
            }
            
            // Configurar o modal com conteúdo customizado
            const modalTitle = modal.querySelector('#modal-title');
            const modalMessage = modal.querySelector('#modal-message');
            
            if (modalTitle) modalTitle.textContent = 'Finalizar Rascunho';
            if (modalMessage) modalMessage.textContent = `Deseja finalizar o rascunho ${rascunhoId} como um pedido oficial? Esta ação não pode ser desfeita.`;
            
            // Substituir o botão OK padrão por botões customizados
            const modalContainer = modal.querySelector('.bg-white');
            const existingButton = modal.querySelector('#modal-close-btn');
            
            if (existingButton) {
                existingButton.remove();
            }
            
            // Criar container para os botões
            const actionsContainer = document.createElement('div');
            actionsContainer.className = 'flex space-x-4 mt-4';
            actionsContainer.innerHTML = `
                <button id="cancelarFinalizacao" class="flex-1 px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600">Cancelar</button>
                <button id="confirmarFinalizacao" class="flex-1 px-4 py-2 bg-green-600 text-white rounded hover:bg-green-700">Finalizar</button>
            `;
            
            modalContainer.appendChild(actionsContainer);
            
            // Mostrar o modal
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            
            // Eventos dos botões
            document.getElementById('cancelarFinalizacao').addEventListener('click', () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                restaurarModalPadrao(modal);
            });
            
            document.getElementById('confirmarFinalizacao').addEventListener('click', () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                restaurarModalPadrao(modal);
                processarFinalizacaoRascunho(rascunhoId);
            });
        }
        
        function restaurarModalPadrao(modal) {
            // Remove os botões customizados e restaura o botão OK padrão
            const actionsContainer = modal.querySelector('.flex.space-x-4');
            if (actionsContainer) {
                actionsContainer.remove();
            }
            
            const modalContainer = modal.querySelector('.bg-white');
            const okButton = document.createElement('button');
            okButton.id = 'modal-close-btn';
            okButton.className = 'px-6 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700';
            okButton.textContent = 'OK';
            okButton.addEventListener('click', () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
            });
            
            modalContainer.appendChild(okButton);
        }
        
        function processarFinalizacaoRascunho(rascunhoId) {
            showToast('Finalizando rascunho...', 'info');
            
            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'success') {
                        showToast(`Rascunho finalizado com sucesso! Pedido: ${response.numeroPedido}`, 'success', 'Pedido Criado');
                        
                        // Recarregar lista de rascunhos
                        if (typeof setupGerenciarRascunhosScreen === 'function') {
                            setupGerenciarRascunhosScreen();
                        }
                    } else {
                        showToast('Erro ao finalizar rascunho: ' + response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    console.error('Erro ao finalizar rascunho:', error);
                    showToast('Erro de comunicação: ' + error.message, 'error');
                })
                .finalizarRascunho(rascunhoId);
        }
        
        function excluirRascunho(rascunhoId) {
            console.log('🗑️ Excluindo rascunho:', rascunhoId);
            
            const modal = document.getElementById('global-modal');
            if (!modal) {
                console.error('Modal global não encontrado');
                return;
            }
            
            // Configurar o modal com conteúdo customizado
            const modalTitle = modal.querySelector('#modal-title');
            const modalMessage = modal.querySelector('#modal-message');
            
            if (modalTitle) modalTitle.textContent = 'Excluir Rascunho';
            if (modalMessage) modalMessage.textContent = `Tem certeza que deseja excluir o rascunho ${rascunhoId}? Esta ação não pode ser desfeita.`;
            
            // Substituir o botão OK padrão por botões customizados
            const modalContainer = modal.querySelector('.bg-white');
            const existingButton = modal.querySelector('#modal-close-btn');
            
            if (existingButton) {
                existingButton.remove();
            }
            
            // Criar container para os botões
            const actionsContainer = document.createElement('div');
            actionsContainer.className = 'flex space-x-4 mt-4';
            actionsContainer.innerHTML = `
                <button id="cancelarExclusao" class="flex-1 px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600">Cancelar</button>
                <button id="confirmarExclusao" class="flex-1 px-4 py-2 bg-red-600 text-white rounded hover:bg-red-700">Excluir</button>
            `;
            
            modalContainer.appendChild(actionsContainer);
            
            // Mostrar o modal
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            
            // Eventos dos botões
            document.getElementById('cancelarExclusao').addEventListener('click', () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                restaurarModalPadrao(modal);
            });
            
            document.getElementById('confirmarExclusao').addEventListener('click', () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                restaurarModalPadrao(modal);
                processarExclusaoRascunho(rascunhoId);
            });
        }
        
        function processarExclusaoRascunho(rascunhoId) {
            showToast('Excluindo rascunho...', 'info');
            
            google.script.run
                .withSuccessHandler(response => {
                    if (response.status === 'success') {
                        showToast('Rascunho excluído com sucesso!', 'success');
                        
                        // Recarregar lista de rascunhos
                        if (typeof setupGerenciarRascunhosScreen === 'function') {
                            setupGerenciarRascunhosScreen();
                        }
                    } else {
                        showToast('Erro ao excluir rascunho: ' + response.message, 'error');
                    }
                })
                .withFailureHandler(error => {
                    console.error('Erro ao excluir rascunho:', error);
                    showToast('Erro de comunicação: ' + error.message, 'error');
                })
                .excluirRascunho(rascunhoId);
        }
        
        function preencherFormularioComRascunho(dadosRascunho) {
            console.log('📝 Preenchendo formulário com dados do rascunho:', dadosRascunho);
            
            // Função para tentar preencher fornecedor com retry
            function tentarPreencherFornecedor(tentativa = 1, maxTentativas = 10) {
                const fornecedorSelect = document.getElementById('fornecedorPedido');
                
                if (fornecedorSelect && fornecedorSelect.options.length > 1) {
                    // Tentar encontrar o fornecedor exato
                    let encontrou = false;
                    for (let i = 0; i < fornecedorSelect.options.length; i++) {
                        if (fornecedorSelect.options[i].value === dadosRascunho.fornecedor || 
                            fornecedorSelect.options[i].textContent === dadosRascunho.fornecedor) {
                            fornecedorSelect.selectedIndex = i;
                            encontrou = true;
                            console.log('✅ Fornecedor selecionado:', dadosRascunho.fornecedor);
                            
                            // Disparar evento de mudança para atualizar o estado
                            const evento = new Event('change', { bubbles: true });
                            fornecedorSelect.dispatchEvent(evento);
                            
                            break;
                        }
                    }
                    
                    if (!encontrou) {
                        console.warn('⚠️ Fornecedor não encontrado nas opções:', dadosRascunho.fornecedor);
                        console.log('📋 Opções disponíveis:', Array.from(fornecedorSelect.options).map(opt => opt.value));
                    }
                } else if (tentativa < maxTentativas) {
                    console.log(`🔄 Tentativa ${tentativa}: Aguardando carregamento de fornecedores...`);
                    setTimeout(() => tentarPreencherFornecedor(tentativa + 1, maxTentativas), 500);
                } else {
                    console.warn('⚠️ Timeout: Não foi possível carregar fornecedores');
                }
            }
            
            // Função para tentar preencher veículo com retry
            function tentarPreencherVeiculo(tentativa = 1, maxTentativas = 10) {
                const veiculoSelect = document.getElementById('nomeVeiculo');
                
                if (veiculoSelect && veiculoSelect.options.length > 1 && dadosRascunho.nomeVeiculo) {
                    // Tentar encontrar o veículo exato
                    let encontrou = false;
                    for (let i = 0; i < veiculoSelect.options.length; i++) {
                        if (veiculoSelect.options[i].value === dadosRascunho.nomeVeiculo || 
                            veiculoSelect.options[i].textContent === dadosRascunho.nomeVeiculo) {
                            veiculoSelect.selectedIndex = i;
                            encontrou = true;
                            console.log('✅ Veículo selecionado:', dadosRascunho.nomeVeiculo);
                            break;
                        }
                    }
                    
                    if (!encontrou) {
                        console.warn('⚠️ Veículo não encontrado nas opções:', dadosRascunho.nomeVeiculo);
                    }
                } else if (tentativa < maxTentativas && dadosRascunho.nomeVeiculo) {
                    console.log(`🔄 Tentativa ${tentativa}: Aguardando carregamento de veículos...`);
                    setTimeout(() => tentarPreencherVeiculo(tentativa + 1, maxTentativas), 500);
                }
            }
            
            // Preencher campos básicos
            if (dadosRascunho.data) {
                const dataPedido = document.getElementById('dataPedido');
                if (dataPedido) {
                    dataPedido.value = dadosRascunho.data;
                    console.log('✅ Data preenchida:', dadosRascunho.data);
                } else {
                    console.warn('⚠️ Campo dataPedido não encontrado');
                }
            }
            
            // Tentar preencher fornecedor com retry
            if (dadosRascunho.fornecedor) {
                tentarPreencherFornecedor();
            }
            
            // Tentar preencher veículo com retry
            if (dadosRascunho.nomeVeiculo) {
                tentarPreencherVeiculo();
            }
            
            if (dadosRascunho.placaVeiculo) {
                const placaInput = document.getElementById('placaVeiculo');
                if (placaInput) {
                    placaInput.value = dadosRascunho.placaVeiculo;
                    console.log('✅ Placa preenchida:', dadosRascunho.placaVeiculo);
                } else {
                    console.warn('⚠️ Campo placaVeiculo não encontrado');
                }
            }
            
            if (dadosRascunho.observacoes) {
                const observacoesInput = document.getElementById('observacoesPedido');
                if (observacoesInput) {
                    observacoesInput.value = dadosRascunho.observacoes;
                    console.log('✅ Observações preenchidas:', dadosRascunho.observacoes);
                } else {
                    console.warn('⚠️ Campo observacoesPedido não encontrado');
                }
            }
            
            // Preencher itens
            if (dadosRascunho.itens && Array.isArray(dadosRascunho.itens)) {
                console.log('📦 Iniciando preenchimento de itens. Total:', dadosRascunho.itens.length);
                
                // Limpar itens existentes
                const tableBody = document.getElementById('itensTableBody');
                if (tableBody) {
                    tableBody.innerHTML = '';
                    itemCounter = 0;
                    console.log('🗑️ Tabela de itens limpa');
                    
                    // Reinicializar array de itens
                    if (!window.itensAdicionados) {
                        window.itensAdicionados = [];
                    }
                    window.itensAdicionados = [];
                    
                } else {
                    console.error('❌ Tabela de itens não encontrada!');
                    return;
                }
                
                // Adicionar itens do rascunho diretamente à tabela
                dadosRascunho.itens.forEach((item, index) => {
                    console.log(`📦 Adicionando item ${index + 1} à tabela:`, item);
                    
                    // Incrementar contador
                    itemCounter++;
                    
                    // Calcular subtotal
                    const quantidade = parseFloat(item.quantidade) || 1;
                    const precoUnitario = parseFloat(item.precoUnitario) || 0;
                    const subtotal = quantidade * precoUnitario;
                    
                    // Criar nova linha na tabela
                    const row = document.createElement('tr');
                    row.setAttribute('data-item-id', itemCounter);
                    row.className = 'hover:bg-gray-50';
                    row.innerHTML = `
                        <td class="px-6 py-4 text-sm text-gray-900">${item.descricao || ''}</td>
                        <td class="px-6 py-4 text-sm text-gray-900 text-center">${quantidade.toLocaleString('pt-BR')}</td>
                        <td class="px-6 py-4 text-sm text-gray-900 text-center">${item.unidade || ''}</td>
                        <td class="px-6 py-4 text-sm text-gray-900 text-right">R$ ${precoUnitario.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                        <td class="px-6 py-4 text-sm font-medium text-gray-900 text-right">R$ ${subtotal.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                        <td class="px-6 py-4 text-center">
                            <button type="button" class="text-red-600 hover:text-red-800 font-medium" onclick="removerItem(${itemCounter})">
                                <i class="fas fa-trash"></i>
                            </button>
                        </td>
                    `;
                    
                    // Adicionar à tabela
                    tableBody.appendChild(row);
                    
                    // Armazenar dados do item
                    window.itensAdicionados.push({
                        id: itemCounter,
                        descricao: item.descricao || '',
                        quantidade: quantidade,
                        unidade: item.unidade || '',
                        precoUnitario: precoUnitario,
                        subtotal: subtotal // Usando 'subtotal' para compatibilidade com calcularTotalGeral()
                    });
                    
                    console.log(`✅ Item ${itemCounter} adicionado à tabela: ${item.descricao}`);
                });
                
                // Recalcular total geral
                setTimeout(() => {
                    console.log('🧮 Recalculando total geral...');
                    calcularTotalGeral();
                }, 200);
            } else {
                console.warn('⚠️ Nenhum item encontrado no rascunho ou itens não é um array');
            }
        }
        
        // Função para limpar dados de edição de rascunho
        function limparDadosEdicaoRascunho() {
            localStorage.removeItem('editandoRascunho');
            localStorage.removeItem('rascunhoId');
            localStorage.removeItem('dadosRascunho');
            console.log('🧹 Dados de edição de rascunho limpos');
        }
        
        function mostrarIndicadorRascunho(rascunhoId) {
            const draftStatus = document.getElementById('draft-status');
            if (draftStatus) {
                draftStatus.classList.remove('hidden');
                draftStatus.innerHTML = `<i class="fas fa-edit mr-2"></i>Editando Rascunho - ID: ${rascunhoId}`;
            }
        }

        // ===============================================
        // FUNÇÃO DE TESTE PARA VERIFICAR ESTADOS
        // ===============================================
        
        window.testarEstadosFornecedores = function() {
            console.log('🔍 === TESTE DE ESTADOS DOS FORNECEDORES ===');
            google.script.run
                .withSuccessHandler(fornecedores => {
                    console.log('📋 Total de fornecedores:', fornecedores.length);
                    fornecedores.forEach((f, index) => {
                        console.log(`${index + 1}. ${f.razao} - Estado: "${f.estado || 'VAZIO'}"`);
                    });
                    
                    const comEstado = fornecedores.filter(f => f.estado && f.estado.trim());
                    console.log(`✅ Fornecedores COM estado: ${comEstado.length}`);
                    console.log(`❌ Fornecedores SEM estado: ${fornecedores.length - comEstado.length}`);
                    
                    if (comEstado.length > 0) {
                        console.log('🎯 Fornecedores que podem ser usados para teste:');
                        comEstado.slice(0, 3).forEach(f => {
                            console.log(`   - ${f.razao} (${f.estado})`);
                        });
                    }
                })
                .withFailureHandler(err => {
                    console.error('❌ Erro ao buscar fornecedores:', err);
                })
                .getFornecedoresList();
        };
        
        window.testarCaptura = function(nomeFornecedor) {
            console.log('🔍 === TESTE DE CAPTURA DE ESTADO ===');
            console.log(`📋 Testando fornecedor: ${nomeFornecedor}`);
            google.script.run
                .withSuccessHandler(resultado => {
                    console.log('✅ Resultado do teste:', resultado);
                })
                .withFailureHandler(err => {
                    console.error('❌ Erro no teste:', err);
                })
                .testarEstadoFornecedor(nomeFornecedor);
        };
        
        // ===============================================
        // INICIALIZAÇÃO DO APP
        // ===============================================
        
        // Expor funções de teste globalmente
        window.testarBotaoAdicionarItem = testarBotaoAdicionarItem;
        window.testarEnterKeydown = testarEnterKeydown;
        window.itemCounter = itemCounter;
        
        document.addEventListener('DOMContentLoaded', () => {
            setupLoginScreen();
            setupCadastroScreen();
            showScreen('screen-login');
        });
    </script>
</body>
</html>