import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';

const NOMES_ITENS = [
  'Cabo ATX',
  'Cabo Bipolar',
  'Cabo Tripolar',
  'Cabo HDMI',
  'Cabo VGA',
  'Memoria RAM',
  'Placa de Rede',
  'Cabo Displayport',
  'Adaptador HDMI',
  'Mouse',
  'Chip Salvy',
  'Alexa',
  'Fonte Zebra',
  'HD',
  'Fonte Avulsa',
  'Fonte Tanca',
  'Cabo de Dados USB',
  'Bateria Notebook',
  'Teclado',
  'Organizador de Cabos',
  'Monitor',
  'Automatizador Ar-Condicionado',
  'Celular',
  'Celular Bloqueado',
  'Impressora Termica',
  'CPU',
  'Coletor Honeywell',
];

const CATEGORIAS = [
  'Limpeza', 'HDMI', 'VGA', 'Memória RAM / Placa de rede', 'ATX',
  'Display Port', 'Adaptador HDMI', 'Mouse', 'Chip Salvy', 'Alexa',
  'Fonte Zebra', 'HDs', 'Fontes Avulsa', 'Fonte Tanca', 'Cabos de dados USB',
  'Bateria de notebook', 'Teclado', 'Organizador de cabos', 'Monitor',
  'Automação ar condicionado', 'Celular bloqueado', 'Impressora térmica',
  'CPU', 'Coletor Honeywell',
];

const STATUS = ['Disponível', 'Em manutenção'];
const STATUS_COLOR = {
  'Disponível': { bg: '#EAF3DE', text: '#27500A' },
  'Em manutenção': { bg: '#FAEEDA', text: '#633806' },
  'Baixado': { bg: '#FCEBEB', text: '#791F1F' },
};

const OPERATIONAL_TABS = {
  disponiveis: {
    title: 'Itens disponíveis',
    shortTitle: 'Disponíveis',
    status: 'Disponível',
    actions: ['Retirado do Estoque', 'Envio para manutenção', 'Item quebrado/danificado'],
    empty: 'Nenhum item disponível no estoque.',
  },
  manutencao: {
    title: 'Em manutenção',
    shortTitle: 'Manutenção',
    status: 'Em manutenção',
    actions: ['Retorno da manutenção', 'Item quebrado/danificado'],
    empty: 'Nenhum item em manutenção.',
  },
};

const ACTION_CONFIG = {
  'Retirado do Estoque': { targetStatus: null, removeFromStock: true, destinationLabel: 'Fora do estoque' },
  'Envio para manutenção': { targetStatus: 'Em manutenção', destinationLabel: 'Em manutenção' },
  'Retorno da manutenção': { targetStatus: 'Disponível', destinationLabel: 'Disponível' },
  'Item quebrado/danificado': { targetStatus: null, removeFromStock: true, destinationLabel: 'Baixado por quebra/dano' },
};

const ACCENT = '#185FA5';
const ACCENT_DARK = '#042C53';

const CLIENT_ID = '37ff5e3e-1558-4add-b4e9-8e5c97e21943';
const TENANT_ID = '9cc5e35a-c64c-4450-ae6d-9bf065f73c61';
const EXCEL_FILE_WEB_URL = 'https://flashcouriercombr.sharepoint.com/:x:/r/sites/Suporte_Tcnico/Documentos%20Partilhados/ESTOQUE%20TI/Estoque%20TI.xlsx?d=w8665de340bf445ac826b93bcaae1bb16&csf=1&web=1&e=FONu2c';
const SHAREPOINT_HOSTNAME = 'flashcouriercombr.sharepoint.com';
const SHAREPOINT_SITE_PATH = '/sites/Suporte_Tcnico';
const EXCEL_FILE_PATH = 'ESTOQUE TI/Estoque TI.xlsx';
const GRAPH_SCOPES = [
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
  'User.Read',
];
const AUTH_REDIRECT_URI = `${window.location.origin}/redirect.html`;

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: AUTH_REDIRECT_URI,
    postLogoutRedirectUri: AUTH_REDIRECT_URI,
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
};

const msalInstance = new PublicClientApplication(msalConfig);

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function fmtDate(d) {
  if (!d) return '—';
  try { return new Date(d).toLocaleDateString('pt-BR'); } catch { return d; }
}
function fmtDateTime(d) {
  if (!d) return '—';
  try { return new Date(d).toLocaleString('pt-BR'); } catch { return d; }
}
function toPositiveInt(value, fallback = 1) {
  const n = parseInt(value, 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}
function getItemQuantity(item) {
  return toPositiveInt(item?.quantidade, 1);
}
function shouldClearResponsavelByAction(action) {
  return Boolean(ACTION_CONFIG[action]?.clearResponsavel);
}
function buildMovementObs(qtd, obs, origem, destino) {
  const cleanObs = (obs || '').trim();
  const partes = [`Qtd. movimentada: ${qtd}`];
  if (origem) partes.push(`Origem: ${origem}`);
  if (destino) partes.push(`Destino: ${destino}`);
  if (cleanObs) partes.push(cleanObs);
  return partes.join(' | ');
}
function parseMovementQuantity(obs) {
  const match = String(obs || '').match(/Qtd\. movimentada:\s*(\d+)/i);
  return match ? match[1] : '—';
}
function rowToItem(values) {
  const [ID, Nome, Categoria, Quantidade, Patrimonio, Status, Responsavel, Localizacao, DataEntrada, Observacoes, CreatedAt] = values;
  return { id: ID, nome: Nome, categoria: Categoria, quantidade: Quantidade || 1, patrimonio: Patrimonio, status: Status, responsavel: Responsavel, localizacao: Localizacao, dataEntrada: DataEntrada, obs: Observacoes, createdAt: CreatedAt };
}
function itemToRow(it) {
  return [it.id, it.nome, it.categoria, it.quantidade || 1, it.patrimonio || '', it.status, it.responsavel || '', it.localizacao || '', it.dataEntrada || '', it.obs || '', it.createdAt || new Date().toISOString()];
}
function rowToMov(values) {
  const [ID, ItemId, ItemNome, Acao, Responsavel, RegistradoPor, Data, Obs] = values;
  return { id: ID, itemId: ItemId, itemName: ItemNome, action: Acao, responsavel: Responsavel, registradoPor: RegistradoPor, date: Data, obs: Obs };
}
function movToRow(mv) {
  return [mv.id, mv.itemId, mv.itemName, mv.action, mv.responsavel || '—', mv.registradoPor, mv.date, mv.obs || ''];
}

function normalizeKey(value) {
  return String(value || '').trim().toLowerCase();
}
function groupKeyForItem(it) {
  const status = it.status || 'Disponível';
  return [
    it.nome,
    it.categoria,
    it.patrimonio,
    status,
    it.localizacao,
    it.obs,
  ].map(normalizeKey).join('||');
}
function groupItemsByQuantity(rows) {
  const map = new Map();
  rows.forEach((it) => {
    const key = groupKeyForItem(it);
    const current = map.get(key);
    if (!current) {
      map.set(key, {
        ...it,
        id: `group_${key}`,
        groupKey: key,
        itemIds: [it.id],
        quantidade: getItemQuantity(it),
        sourceCount: 1,
        responsavel: '',
      });
    } else {
      current.itemIds.push(it.id);
      current.quantidade += getItemQuantity(it);
      current.sourceCount += 1;
    }
  });
  return Array.from(map.values()).sort((a, b) => String(a.nome || '').localeCompare(String(b.nome || ''), 'pt-BR'));
}
function statusDestinoPorAcao(action) {
  const cfg = ACTION_CONFIG[action];
  if (!cfg) return '';
  return cfg.destinationLabel || (cfg.removeFromStock ? 'Removido do estoque' : cfg.targetStatus);
}

export default function InventarioTI() {
  const [msalReady, setMsalReady] = useState(false);
  const [account, setAccount] = useState(null);
  const [authError, setAuthError] = useState('');
  const [connecting, setConnecting] = useState(false);
  const graphCtx = useRef({ driveId: null, itemId: null });

  const [items, setItems] = useState([]);
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [tab, setTab] = useState('disponiveis');

  const [search, setSearch] = useState('');
  const [filterCat, setFilterCat] = useState('Todas');

  const [modalOpen, setModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [moveModalOpen, setMoveModalOpen] = useState(false);
  const [moveItem, setMoveItem] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    msalInstance.initialize().then(async () => {
      const redirectResult = await msalInstance.handleRedirectPromise().catch(() => null);
      const accs = msalInstance.getAllAccounts();
      const activeAccount = redirectResult?.account || msalInstance.getActiveAccount() || accs[0] || null;
      if (activeAccount) {
        msalInstance.setActiveAccount(activeAccount);
        setAccount(activeAccount);
      }
      setMsalReady(true);
    }).catch(() => {
      setAuthError('Falha ao inicializar autenticação Microsoft.');
      setMsalReady(true);
    });
  }, []);

  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast(null), 3000);
  }

  async function getToken() {
    if (!account) throw new Error('not signed in');
    const tokenRequest = { scopes: GRAPH_SCOPES, account, redirectUri: AUTH_REDIRECT_URI };
    try {
      const res = await msalInstance.acquireTokenSilent(tokenRequest);
      return res.accessToken;
    } catch (e) {
      const precisaInteracao = e instanceof InteractionRequiredAuthError
        || ['interaction_required', 'consent_required', 'login_required'].includes(e?.errorCode);

      if (!precisaInteracao) throw e;

      const res = await msalInstance.acquireTokenPopup(tokenRequest);
      return res.accessToken;
    }
  }

  async function graphFetch(path, options = {}) {
    const token = await getToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(options.headers || {}),
      },
    });
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new Error(`Graph ${res.status}: ${text.slice(0, 300)}`);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  async function login() {
    setAuthError('');
    setConnecting(true);
    try {
      const res = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES, redirectUri: AUTH_REDIRECT_URI });
      msalInstance.setActiveAccount(res.account);
      setAccount(res.account);
    } catch (e) {
      setAuthError('Falha no login: ' + (e.message || 'verifique o Redirect URI no Azure AD.'));
    } finally {
      setConnecting(false);
    }
  }

  async function logout() {
    try { await msalInstance.logoutPopup(); } catch {}
    setAccount(null);
    setItems([]);
    setMovements([]);
    graphCtx.current = { driveId: null, itemId: null };
  }

  function encodeGraphPath(path) {
    return path
      .split('/')
      .filter(Boolean)
      .map((part) => encodeURIComponent(part))
      .join('/');
  }

  async function tryGraphResolve(label, path, errors) {
    try {
      const item = await graphFetch(path);
      if (item?.id) return item;
      errors.push(`${label}: retorno sem id`);
      return null;
    } catch (e) {
      errors.push(`${label}: ${e.message}`);
      return null;
    }
  }

  const ensureGraphContext = useCallback(async () => {
    if (graphCtx.current.driveId && graphCtx.current.itemId) return graphCtx.current;

    const errors = [];
    const encodedFilePath = encodeGraphPath(EXCEL_FILE_PATH);
    const fallbackPaths = [
      encodedFilePath,
      encodeGraphPath(`Documentos Partilhados/${EXCEL_FILE_PATH}`),
      encodeGraphPath(`Shared Documents/${EXCEL_FILE_PATH}`),
      encodeGraphPath(`Documents/${EXCEL_FILE_PATH}`),
    ];

    const site = await graphFetch(`/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE_PATH}`);
    if (!site?.id) throw new Error('não foi possível localizar o site do SharePoint');

    // 1) Tenta primeiro pela biblioteca padrão do site.
    for (const filePath of fallbackPaths) {
      const item = await tryGraphResolve(
        `drive padrão: ${decodeURIComponent(filePath)}`,
        `/sites/${site.id}/drive/root:/${filePath}`,
        errors,
      );
      if (item?.id) {
        graphCtx.current = {
          driveId: item.parentReference?.driveId,
          itemId: item.id,
        };
        if (graphCtx.current.driveId && graphCtx.current.itemId) return graphCtx.current;
      }
    }

    // 2) Se a biblioteca padrão não for a correta, percorre todas as bibliotecas/document libraries do site.
    const drivesRes = await graphFetch(`/sites/${site.id}/drives`);
    const drives = drivesRes?.value || [];

    for (const drive of drives) {
      for (const filePath of fallbackPaths) {
        const item = await tryGraphResolve(
          `${drive.name || drive.id}: ${decodeURIComponent(filePath)}`,
          `/drives/${drive.id}/root:/${filePath}`,
          errors,
        );
        if (item?.id) {
          graphCtx.current = {
            driveId: item.parentReference?.driveId || drive.id,
            itemId: item.id,
          };
          return graphCtx.current;
        }
      }
    }

    throw new Error(
      `arquivo não encontrado no SharePoint. Site: ${SHAREPOINT_HOSTNAME}${SHAREPOINT_SITE_PATH}. Caminho esperado: ${EXCEL_FILE_PATH}. Tentativas: ${errors.slice(0, 6).join(' | ')}`,
    );
  }, [account]);

  function workbookBase(driveId, itemId) {
    return `/drives/${driveId}/items/${itemId}/workbook`;
  }

  async function loadAllData() {
    setLoading(true);
    try {
      const { driveId, itemId } = await ensureGraphContext();
      const base = workbookBase(driveId, itemId);
      const [itemsRes, movsRes] = await Promise.all([
        graphFetch(`${base}/tables('Itens')/rows`),
        graphFetch(`${base}/tables('Movimentacoes')/rows`),
      ]);
      const itemRows = (itemsRes.value || []).filter((r) => {
        const id = r?.values?.[0]?.[0];
        return id && !String(id).startsWith('__MODELO__');
      });
      const movRows = (movsRes.value || []).filter((r) => {
        const id = r?.values?.[0]?.[0];
        return id && !String(id).startsWith('__MODELO__');
      });
      setItems(itemRows.map((r) => rowToItem(r.values[0])).reverse());
      setMovements(movRows.map((r) => rowToMov(r.values[0])).reverse());
    } catch (e) {
      showToast('Erro ao carregar planilha: ' + e.message);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (account) loadAllData();
  }, [account]);

  async function addRow(tableName, rowValues) {
    const { driveId, itemId } = await ensureGraphContext();
    await graphFetch(`${workbookBase(driveId, itemId)}/tables('${tableName}')/rows`, {
      method: 'POST',
      body: JSON.stringify({ values: [rowValues] }),
    });
  }

  async function findRowIndex(tableName, idValue) {
    const { driveId, itemId } = await ensureGraphContext();
    const res = await graphFetch(`${workbookBase(driveId, itemId)}/tables('${tableName}')/rows`);
    return (res.value || []).findIndex((r) => r.values[0][0] === idValue);
  }

  async function updateRowById(tableName, idValue, rowValues) {
    const idx = await findRowIndex(tableName, idValue);
    if (idx === -1) throw new Error('linha não encontrada');
    const { driveId, itemId } = await ensureGraphContext();
    await graphFetch(`${workbookBase(driveId, itemId)}/tables('${tableName}')/rows/itemAt(index=${idx})`, {
      method: 'PATCH',
      body: JSON.stringify({ values: [rowValues] }),
    });
  }

  async function deleteRowById(tableName, idValue) {
    const idx = await findRowIndex(tableName, idValue);
    if (idx === -1) return;
    const { driveId, itemId } = await ensureGraphContext();
    await graphFetch(`${workbookBase(driveId, itemId)}/tables('${tableName}')/rows/itemAt(index=${idx})`, { method: 'DELETE' });
  }

  function openNewModal() {
    setEditingItem({ id: null, nome: NOMES_ITENS[0], categoria: CATEGORIAS[0], patrimonio: '', status: 'Disponível', quantidade: 1, responsavel: '', localizacao: '', dataEntrada: new Date().toISOString().slice(0, 10), obs: '' });
    setModalOpen(true);
  }
  function openAddSimilarModal(itemGroup) {
    const cfg = OPERATIONAL_TABS[tab] || OPERATIONAL_TABS.disponiveis;
    setEditingItem({
      id: null,
      nome: itemGroup.nome || NOMES_ITENS[0],
      categoria: itemGroup.categoria || CATEGORIAS[0],
      patrimonio: '',
      status: cfg.status || 'Disponível',
      quantidade: 1,
      responsavel: '',
      localizacao: itemGroup.localizacao || '',
      dataEntrada: new Date().toISOString().slice(0, 10),
      obs: '',
    });
    setModalOpen(true);
  }
  function openEditModal(item) {
    setEditingItem({ quantidade: 1, ...item, nome: NOMES_ITENS.includes(item?.nome) ? item.nome : NOMES_ITENS[0], responsavel: '' });
    setModalOpen(true);
  }

  async function saveItem(e) {
    e.preventDefault();
    if (!editingItem.nome.trim()) { showToast('Informe o nome do item.'); return; }
    setSyncing(true);
    try {
      const isNew = !editingItem.id;
      const savedItem = { ...editingItem, responsavel: '', quantidade: Math.max(1, parseInt(editingItem.quantidade, 10) || 1), id: isNew ? uid() : editingItem.id, createdAt: isNew ? new Date().toISOString() : editingItem.createdAt };
      if (isNew) {
        await addRow('Itens', itemToRow(savedItem));
        setItems((prev) => [savedItem, ...prev]);
      } else {
        await updateRowById('Itens', savedItem.id, itemToRow(savedItem));
        setItems((prev) => prev.map((it) => (it.id === savedItem.id ? savedItem : it)));
      }
      const mv = { id: uid(), itemId: savedItem.id, itemName: savedItem.nome, action: isNew ? 'Cadastro' : 'Edição', responsavel: '—', registradoPor: account.name || account.username, date: new Date().toISOString(), obs: isNew ? 'Item cadastrado' : 'Dados atualizados' };
      await addRow('Movimentacoes', movToRow(mv));
      setMovements((prev) => [mv, ...prev]);
      setModalOpen(false); setEditingItem(null);
      showToast(isNew ? 'Item cadastrado.' : 'Item atualizado.');
    } catch (e) {
      showToast('Erro: ' + e.message);
    } finally { setSyncing(false); }
  }

  async function deleteItem(id) {
    const it = items.find((x) => x.id === id);
    setSyncing(true);
    try {
      await deleteRowById('Itens', id);
      setItems((prev) => prev.filter((x) => x.id !== id));
      setConfirmDelete(null);
      const mv = { id: uid(), itemId: id, itemName: it?.nome, action: 'Remoção', responsavel: '—', registradoPor: account.name || account.username, date: new Date().toISOString(), obs: 'Item removido' };
      await addRow('Movimentacoes', movToRow(mv));
      setMovements((prev) => [mv, ...prev]);
      showToast('Item removido.');
    } catch (e) { showToast('Erro: ' + e.message); }
    finally { setSyncing(false); }
  }

  function openMoveModal(itemGroup) {
    const cfg = OPERATIONAL_TABS[tab] || OPERATIONAL_TABS.disponiveis;
    const actions = cfg.actions;
    const firstAction = actions[0];
    const qtdAtual = getItemQuantity(itemGroup);
    setMoveItem({
      sourceItemIds: itemGroup.itemIds || [itemGroup.id],
      itemName: itemGroup.nome,
      sourceStatus: cfg.status,
      action: firstAction,
      allowedActions: actions,
      responsavel: '',
      obs: '',
      quantidade: 1,
      quantidadeDisponivel: qtdAtual,
    });
    setMoveModalOpen(true);
  }

  async function saveMovement(e) {
    e.preventDefault();
    const cfg = ACTION_CONFIG[moveItem.action];
    if (!cfg) { showToast('Ação inválida.'); return; }

    const sourceIds = moveItem.sourceItemIds || [moveItem.itemId];
    const currentSources = items
      .filter((it) => sourceIds.includes(it.id))
      .filter((it) => (it.status || 'Disponível') === moveItem.sourceStatus);

    if (currentSources.length === 0) { showToast('Item não encontrado para movimentação.'); return; }

    const qtdAtual = currentSources.reduce((sum, it) => sum + getItemQuantity(it), 0);
    const qtdMovimentada = toPositiveInt(moveItem.quantidade, 1);

    if (qtdMovimentada > qtdAtual) {
      showToast(`Quantidade inválida. Esse grupo possui apenas ${qtdAtual} unidade(s).`);
      return;
    }


    setSyncing(true);
    try {
      const responsavelFinal = '';
      const destino = statusDestinoPorAcao(moveItem.action);
      let restanteParaMover = qtdMovimentada;
      const localItems = [...items];

      for (const source of currentSources) {
        if (restanteParaMover <= 0) break;

        const idxLocal = localItems.findIndex((it) => it.id === source.id);
        if (idxLocal === -1) continue;

        const qtdSource = getItemQuantity(source);
        const qtdConsumida = Math.min(qtdSource, restanteParaMover);
        const qtdRestante = qtdSource - qtdConsumida;
        restanteParaMover -= qtdConsumida;

        if (cfg.removeFromStock) {
          if (qtdRestante > 0) {
            const remainingItem = { ...source, quantidade: qtdRestante };
            await updateRowById('Itens', remainingItem.id, itemToRow(remainingItem));
            localItems[idxLocal] = remainingItem;
          } else {
            await deleteRowById('Itens', source.id);
            localItems.splice(idxLocal, 1);
          }
          continue;
        }

        if (qtdConsumida === qtdSource) {
          const updated = {
            ...source,
            status: cfg.targetStatus,
            responsavel: responsavelFinal,
          };
          await updateRowById('Itens', updated.id, itemToRow(updated));
          localItems[idxLocal] = updated;
        } else {
          const remainingItem = { ...source, quantidade: qtdRestante };
          const movedItem = {
            ...source,
            id: uid(),
            quantidade: qtdConsumida,
            status: cfg.targetStatus,
            responsavel: responsavelFinal,
            createdAt: new Date().toISOString(),
          };

          await updateRowById('Itens', remainingItem.id, itemToRow(remainingItem));
          await addRow('Itens', itemToRow(movedItem));

          localItems[idxLocal] = remainingItem;
          localItems.unshift(movedItem);
        }
      }

      const mv = {
        id: uid(),
        itemId: sourceIds.join(','),
        itemName: moveItem.itemName,
        action: moveItem.action,
        responsavel: responsavelFinal || '—',
        registradoPor: account.name || account.username,
        date: new Date().toISOString(),
        obs: buildMovementObs(qtdMovimentada, moveItem.obs, moveItem.sourceStatus, destino),
      };

      await addRow('Movimentacoes', movToRow(mv));
      setItems(localItems);
      setMovements((prev) => [mv, ...prev]);
      setMoveModalOpen(false); setMoveItem(null);
      showToast('Movimentação registrada.');
    } catch (e) { showToast('Erro: ' + e.message); }
    finally { setSyncing(false); }
  }

  const matchSearchAndCategory = useCallback((it) => {
    const s = search.toLowerCase();
    const matchSearch = !search
      || String(it.nome || '').toLowerCase().includes(s)
      || String(it.patrimonio || '').toLowerCase().includes(s)
      || String(it.localizacao || '').toLowerCase().includes(s);
    return matchSearch && (filterCat === 'Todas' || it.categoria === filterCat);
  }, [search, filterCat]);

  const currentScreen = OPERATIONAL_TABS[tab] || null;

  const filteredGroups = useMemo(() => {
    if (!currentScreen) return [];
    const rows = items
      .filter((it) => (it.status || 'Disponível') === currentScreen.status)
      .filter(matchSearchAndCategory);
    return groupItemsByQuantity(rows);
  }, [items, currentScreen, matchSearchAndCategory]);

  const stats = useMemo(() => {
    const qtdPorStatus = (status) => items
      .filter((i) => (i.status || 'Disponível') === status)
      .reduce((sum, i) => sum + getItemQuantity(i), 0);

    const disp = qtdPorStatus('Disponível');
    const manut = qtdPorStatus('Em manutenção');

    return {
      totalQtd: disp + manut,
      disp,
      manut,
    };
  }, [items]);

  const qtdPorCategoria = useMemo(() => {
    const map = new Map();
    items.forEach((it) => {
      const status = it.status || 'Disponível';
      if (!['Disponível', 'Em manutenção'].includes(status)) return;
      const key = `${it.categoria || 'Sem categoria'}||${it.nome || 'Sem nome'}`;
      if (!map.has(key)) {
        map.set(key, {
          categoria: it.categoria || 'Sem categoria',
          nome: it.nome || 'Sem nome',
          disponivel: 0,
          manutencao: 0,
          total: 0,
        });
      }
      const row = map.get(key);
      const qtd = getItemQuantity(it);
      if (status === 'Disponível') row.disponivel += qtd;
      if (status === 'Em manutenção') row.manutencao += qtd;
      row.total += qtd;
    });
    return Array.from(map.values()).sort((a, b) => {
      const c = a.categoria.localeCompare(b.categoria, 'pt-BR');
      return c || a.nome.localeCompare(b.nome, 'pt-BR');
    });
  }, [items]);

  if (!msalReady) return <div className="min-h-screen flex items-center justify-center bg-stone-100"><p className="text-stone-500 text-sm">Carregando…</p></div>;

  if (!account) return (
    <div className="min-h-screen flex items-center justify-center bg-stone-100 px-4">
      <div className="bg-white rounded-lg border border-stone-200 p-6 w-full max-w-sm text-center">
        <p className="text-xs tracking-widest uppercase text-stone-400 mb-1">Flash Courier · TI</p>
        <h1 className="text-lg font-semibold text-stone-900 mb-1">Inventário de equipamentos</h1>
        <p className="text-sm text-stone-500 mb-5">Entre com sua conta Microsoft 365 para acessar a planilha "Estoque TI".</p>
        <button onClick={login} disabled={connecting} className="w-full px-4 py-2 text-sm rounded font-medium disabled:opacity-60" style={{ backgroundColor: ACCENT, color: '#fff' }}>
          {connecting ? 'Conectando…' : 'Entrar com Microsoft'}
        </button>
        {authError && <p className="text-xs text-red-700 mt-3 text-left">{authError}</p>}
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-stone-100 text-stone-900 font-sans">
      <header className="bg-stone-900 text-stone-100 px-6 py-5 sticky top-0 z-20 border-b-4" style={{ borderColor: ACCENT }}>
        <div className="max-w-6xl mx-auto flex items-center justify-between flex-wrap gap-3">
          <div>
            <p className="text-xs tracking-widest uppercase text-stone-400">Flash Courier · TI</p>
            <h1 className="text-xl font-semibold">Sala de equipamentos — inventário</h1>
          </div>
          <div className="flex items-center gap-3">
            {syncing && <span className="text-xs text-stone-400">Sincronizando…</span>}
            <button onClick={logout} className="text-xs text-stone-400 hover:text-stone-200 underline">{account.name || account.username}</button>
            <button onClick={loadAllData} className="px-3 py-2 text-sm rounded border border-stone-600 hover:bg-stone-800 transition">Atualizar</button>
            <button onClick={openNewModal} className="px-3 py-2 text-sm rounded font-medium" style={{ backgroundColor: ACCENT, color: '#fff' }}>+ Novo item</button>
          </div>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-6 pt-6">
        {loading ? <p className="text-sm text-stone-400 py-10 text-center">Carregando dados da planilha…</p> : (
          <>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
              <StatCard label="Itens disponíveis" value={stats.disp} color={STATUS_COLOR['Disponível']} />
              <StatCard label="Em manutenção" value={stats.manut} color={STATUS_COLOR['Em manutenção']} />
              <StatCard label="Quantidade total" value={stats.totalQtd} color={{ text: ACCENT_DARK }} />
            </div>

            <div className="flex gap-1 mb-4 border-b border-stone-300 overflow-x-auto">
              <TabButton active={tab === 'disponiveis'} onClick={() => setTab('disponiveis')}>Itens disponíveis</TabButton>
              <TabButton active={tab === 'manutencao'} onClick={() => setTab('manutencao')}>Em manutenção</TabButton>
              <TabButton active={tab === 'quantidades'} onClick={() => setTab('quantidades')}>Quantidade por categoria</TabButton>
              <TabButton active={tab === 'historico'} onClick={() => setTab('historico')}>Histórico</TabButton>
            </div>

            {currentScreen && (
              <>
                <div className="flex flex-wrap gap-2 mb-4">
                  <input type="text" placeholder="Buscar por nome, patrimônio ou local" value={search} onChange={(e) => setSearch(e.target.value)} className="flex-1 min-w-[220px] px-3 py-2 rounded border border-stone-300 bg-white text-sm" />
                  <select value={filterCat} onChange={(e) => setFilterCat(e.target.value)} className="px-3 py-2 rounded border border-stone-300 bg-white text-sm">
                    <option>Todas</option>{CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
                  </select>
                </div>

                <div className="flex items-center justify-between mb-3">
                  <div>
                    <h2 className="font-semibold text-stone-900">{currentScreen.title}</h2>
                    <p className="text-xs text-stone-500">Itens agrupados por descrição/categoria e somados pela quantidade.</p>
                  </div>
                  <span className="text-xs text-stone-500">{filteredGroups.length} grupo(s)</span>
                </div>

                {filteredGroups.length === 0
                  ? <div className="text-center py-16 text-stone-400 border border-dashed border-stone-300 rounded-lg">{items.length === 0 ? 'Nenhum item cadastrado. Clique em "+ Novo item".' : currentScreen.empty}</div>
                  : <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3 pb-10">{filteredGroups.map((it) => <ItemCard key={it.id} item={it} onEdit={it.itemIds?.length === 1 ? () => openEditModal(items.find((row) => row.id === it.itemIds[0]) || it) : null} onMove={() => openMoveModal(it)} onAdd={() => openAddSimilarModal(it)} />)}</div>}
              </>
            )}

            {tab === 'quantidades' && (
              <div className="pb-10">
                {qtdPorCategoria.length === 0
                  ? <div className="text-center py-16 text-stone-400 border border-dashed border-stone-300 rounded-lg">Nenhum item cadastrado.</div>
                  : <div className="bg-white rounded-lg border border-stone-200 overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead className="bg-stone-50 text-stone-500 text-xs uppercase">
                          <tr>
                            <th className="text-left px-4 py-2 font-medium">Categoria</th>
                            <th className="text-left px-4 py-2 font-medium">Item</th>
                            <th className="text-right px-4 py-2 font-medium">Disponível</th>
                            <th className="text-right px-4 py-2 font-medium">Manutenção</th>
                            <th className="text-right px-4 py-2 font-medium">Total</th>
                          </tr>
                        </thead>
                        <tbody>{qtdPorCategoria.map((row) => <tr key={`${row.categoria}-${row.nome}`} className="border-t border-stone-100"><td className="px-4 py-2 text-stone-500">{row.categoria}</td><td className="px-4 py-2 font-medium">{row.nome}</td><td className="px-4 py-2 text-right">{row.disponivel}</td><td className="px-4 py-2 text-right">{row.manutencao}</td><td className="px-4 py-2 text-right font-semibold" style={{ color: ACCENT_DARK }}>{row.total}</td></tr>)}</tbody>
                      </table>
                    </div>}
              </div>
            )}

            {tab === 'historico' && (
              <div className="pb-10">
                {movements.length === 0
                  ? <div className="text-center py-16 text-stone-400 border border-dashed border-stone-300 rounded-lg">Nenhuma movimentação registrada.</div>
                  : <div className="bg-white rounded-lg border border-stone-200 overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead className="bg-stone-50 text-stone-500 text-xs uppercase"><tr><th className="text-left px-4 py-2 font-medium">Data</th><th className="text-left px-4 py-2 font-medium">Item</th><th className="text-left px-4 py-2 font-medium">Ação</th><th className="text-right px-4 py-2 font-medium">Qtd.</th><th className="text-left px-4 py-2 font-medium">Registrado por</th><th className="text-left px-4 py-2 font-medium">Obs / detalhes</th></tr></thead>
                        <tbody>{movements.map((mv) => <tr key={mv.id} className="border-t border-stone-100"><td className="px-4 py-2 whitespace-nowrap text-stone-500">{fmtDateTime(mv.date)}</td><td className="px-4 py-2 font-medium">{mv.itemName}</td><td className="px-4 py-2">{mv.action}</td><td className="px-4 py-2 text-right font-medium">{parseMovementQuantity(mv.obs)}</td><td className="px-4 py-2 text-stone-500">{mv.registradoPor || '—'}</td><td className="px-4 py-2 text-stone-500 min-w-[280px]">{mv.obs || '—'}</td></tr>)}</tbody>
                      </table>
                    </div>}
              </div>
            )}
          </>
        )}
      </div>

      {modalOpen && editingItem && (
        <Modal onClose={() => { setModalOpen(false); setEditingItem(null); }} title={editingItem.id ? 'Editar item' : 'Novo item'}>
          <form onSubmit={saveItem} className="space-y-3">
            <Field label="Nome do item"><select value={editingItem.nome} onChange={(e) => setEditingItem({ ...editingItem, nome: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" required>{NOMES_ITENS.map((nome) => <option key={nome}>{nome}</option>)}</select></Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Categoria / caixa"><select value={editingItem.categoria} onChange={(e) => setEditingItem({ ...editingItem, categoria: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm">{CATEGORIAS.map((c) => <option key={c}>{c}</option>)}</select></Field>
              <Field label="Quantidade"><input type="number" min="1" value={editingItem.quantidade ?? 1} onChange={(e) => setEditingItem({ ...editingItem, quantidade: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
            </div>
            <Field label="Status"><select value={editingItem.status} onChange={(e) => setEditingItem({ ...editingItem, status: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm">{STATUS.map((s) => <option key={s}>{s}</option>)}</select></Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Nº patrimônio / serial"><input type="text" value={editingItem.patrimonio} onChange={(e) => setEditingItem({ ...editingItem, patrimonio: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
              <Field label="Localização"><input type="text" value={editingItem.localizacao} onChange={(e) => setEditingItem({ ...editingItem, localizacao: e.target.value })} placeholder="Ex: Caixa HDMI" className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
            </div>
            <Field label="Data de entrada"><input type="date" value={editingItem.dataEntrada} onChange={(e) => setEditingItem({ ...editingItem, dataEntrada: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
            <Field label="Observações"><textarea value={editingItem.obs} onChange={(e) => setEditingItem({ ...editingItem, obs: e.target.value })} rows={2} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
            <div className="flex justify-end gap-2 pt-2">
              <button type="button" onClick={() => { setModalOpen(false); setEditingItem(null); }} className="px-4 py-2 text-sm rounded border border-stone-300">Cancelar</button>
              <button type="submit" disabled={syncing} className="px-4 py-2 text-sm rounded font-medium disabled:opacity-60" style={{ backgroundColor: ACCENT, color: '#fff' }}>{syncing ? 'Salvando…' : 'Salvar'}</button>
            </div>
          </form>
        </Modal>
      )}

      {moveModalOpen && moveItem && (
        <Modal onClose={() => { setMoveModalOpen(false); setMoveItem(null); }} title={`Movimentar — ${moveItem.itemName}`}>
          <form onSubmit={saveMovement} className="space-y-3">
            <Field label="Ação">
              <select value={moveItem.action} onChange={(e) => setMoveItem({ ...moveItem, action: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm">
                {(moveItem.allowedActions || []).map((action) => <option key={action}>{action}</option>)}
              </select>
            </Field>

            <div className="grid grid-cols-2 gap-3">
              <Field label="Quantidade a movimentar"><input type="number" min="1" max={moveItem.quantidadeDisponivel || 1} value={moveItem.quantidade ?? 1} onChange={(e) => setMoveItem({ ...moveItem, quantidade: e.target.value })} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
              <Field label="Destino"><input type="text" value={statusDestinoPorAcao(moveItem.action)} readOnly className="w-full px-3 py-2 rounded border border-stone-200 bg-stone-50 text-sm text-stone-500" /></Field>
            </div>

            <p className="text-xs text-stone-400 -mt-1">Quantidade neste grupo: {moveItem.quantidadeDisponivel || 1}. Ao movimentar parte da quantidade, o sistema mantém o saldo na tela atual e envia somente a quantidade escolhida para o destino.</p>


            <Field label="Observação"><textarea value={moveItem.obs} onChange={(e) => setMoveItem({ ...moveItem, obs: e.target.value })} rows={2} className="w-full px-3 py-2 rounded border border-stone-300 text-sm" /></Field>
            <p className="text-xs text-stone-400">Registrado por: {account.name || account.username}</p>
            <div className="flex justify-end gap-2 pt-2">
              <button type="button" onClick={() => { setMoveModalOpen(false); setMoveItem(null); }} className="px-4 py-2 text-sm rounded border border-stone-300">Cancelar</button>
              <button type="submit" disabled={syncing} className="px-4 py-2 text-sm rounded font-medium disabled:opacity-60" style={{ backgroundColor: ACCENT, color: '#fff' }}>{syncing ? 'Salvando…' : 'Registrar'}</button>
            </div>
          </form>
        </Modal>
      )}

      {confirmDelete && (
        <Modal onClose={() => setConfirmDelete(null)} title="Remover item">
          <p className="text-sm text-stone-600 mb-4">Tem certeza que deseja remover <strong>{confirmDelete.nome}</strong>? Essa ação não pode ser desfeita.</p>
          <div className="flex justify-end gap-2">
            <button onClick={() => setConfirmDelete(null)} className="px-4 py-2 text-sm rounded border border-stone-300">Cancelar</button>
            <button onClick={() => deleteItem(confirmDelete.id)} disabled={syncing} className="px-4 py-2 text-sm rounded font-medium bg-red-700 text-white disabled:opacity-60">Remover</button>
          </div>
        </Modal>
      )}

      {toast && <div className="fixed bottom-5 left-1/2 -translate-x-1/2 bg-stone-900 text-white text-sm px-4 py-2 rounded shadow-lg z-50 max-w-md text-center">{toast}</div>}
    </div>
  );
}

function StatCard({ label, value, color }) {
  return <div className="bg-white rounded-lg border border-stone-200 px-4 py-3"><p className="text-xs text-stone-500 mb-1">{label}</p><p className="text-2xl font-semibold" style={color ? { color: color.text } : undefined}>{value}</p></div>;
}
function TabButton({ active, onClick, children }) {
  return <button onClick={onClick} className={`px-4 py-2 text-sm font-medium border-b-2 transition ${active ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-400 hover:text-stone-600'}`}>{children}</button>;
}
function ItemCard({ item, onEdit, onMove, onAdd }) {
  const color = STATUS_COLOR[item.status] || STATUS_COLOR['Disponível'];
  return (
    <div className="bg-white rounded-lg border border-stone-200 p-4 flex flex-col gap-2">
      <div className="flex items-start justify-between gap-2">
        <div><p className="font-medium text-stone-900 leading-snug">{item.nome}</p><p className="text-xs text-stone-400">{item.categoria}</p></div>
        <span className="text-xs px-2 py-1 rounded font-medium whitespace-nowrap" style={{ backgroundColor: color.bg, color: color.text }}>{item.status}</span>
      </div>
      <div className="text-xs text-stone-500 space-y-0.5 mt-1">
        <p>Quantidade: <span className="font-medium" style={{ color: ACCENT_DARK }}>{item.quantidade || 1}</span></p>
        {item.sourceCount > 1 && <p>Registros agrupados: {item.sourceCount}</p>}
        {item.patrimonio && <p>Patrimônio: {item.patrimonio}</p>}
        {item.localizacao && <p>Local: {item.localizacao}</p>}
        <p>Entrada: {fmtDate(item.dataEntrada)}</p>
      </div>
      {item.obs && <p className="text-xs text-stone-400 italic border-t border-stone-100 pt-2 mt-1">{item.obs}</p>}
      <div className="flex gap-2 mt-2 pt-2 border-t border-stone-100">
        <button onClick={onMove} className="text-xs px-2 py-1.5 rounded border border-stone-300 hover:bg-stone-50">Movimentar</button>
        <button onClick={onAdd} className="flex-1 text-xs px-2 py-1.5 rounded font-medium" style={{ backgroundColor: ACCENT, color: '#fff' }}>Adicionar</button>
        {onEdit && <button onClick={onEdit} className="text-xs px-2 py-1.5 rounded border border-stone-300 hover:bg-stone-50">Editar</button>}
      </div>
    </div>
  );
}
function Field({ label, children }) {
  return <label className="block"><span className="block text-xs font-medium text-stone-500 mb-1">{label}</span>{children}</label>;
}
function Modal({ title, children, onClose }) {
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-lg shadow-xl max-w-md w-full max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-stone-100">
          <h2 className="font-semibold text-stone-900">{title}</h2>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-700 text-xl leading-none">×</button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}
