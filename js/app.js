/**
 * 通讯录 · 核心逻辑
 * 数据保存在浏览器 localStorage，导入导出使用 SheetJS 读写 Excel
 */
(function () {
  'use strict';

  var STORAGE_KEY = 'contact-book:v1';
  var EXCEL_HEADERS = ['姓名', '是否收藏', '手机号', '邮箱', '社交账号', '住址', '备注'];

  /** 首次打开时的示例数据，方便直接体验功能 */
  var SAMPLE_DATA = [
    { name: '张三', favorite: true, phone: '13800138000', email: 'zhangsan@example.com', social: '微信 zhangsan_wx', address: '福建省福州市福州大学旗山校区', note: '同组队友' },
    { name: '李四', favorite: true, phone: '13900139000', email: 'lisi@example.com', social: 'QQ 10001', address: '福建省厦门市思明区', note: '班长' },
    { name: '王五', favorite: false, phone: '13700137000', email: 'wangwu@example.com', social: '微博 @wangwu', address: '福建省泉州市丰泽区', note: '' }
  ];

  var state = {
    contacts: [],
    keyword: '',
    onlyFavorite: false,
    editingId: null
  };

  var els = {
    list: document.getElementById('contactList'),
    empty: document.getElementById('emptyState'),
    count: document.getElementById('countLabel'),
    search: document.getElementById('searchInput'),
    favFilter: document.getElementById('favFilterBtn'),
    importBtn: document.getElementById('importBtn'),
    exportBtn: document.getElementById('exportBtn'),
    addBtn: document.getElementById('addBtn'),
    fileInput: document.getElementById('fileInput'),
    mask: document.getElementById('mask'),
    form: document.getElementById('contactForm'),
    dialogTitle: document.getElementById('dialogTitle'),
    closeBtn: document.getElementById('closeBtn'),
    cancelBtn: document.getElementById('cancelBtn'),
    toast: document.getElementById('toast')
  };

  /* ================= 数据存取 ================= */

  function loadContacts() {
    var raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      state.contacts = SAMPLE_DATA.map(function (item, index) {
        return Object.assign({ id: createId(index) }, item);
      });
      saveContacts();
      return;
    }
    try {
      var parsed = JSON.parse(raw);
      state.contacts = Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      state.contacts = [];
    }
  }

  function saveContacts() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.contacts));
  }

  var idSeed = 0;

  function createId(seed) {
    idSeed += 1;
    return 'c' + Date.now().toString(36) + '_' + (seed === undefined ? idSeed : seed);
  }

  /* ================= 渲染 ================= */

  /** 收藏优先，其次按姓名拼音排序 */
  function sortContacts(list) {
    return list.slice().sort(function (a, b) {
      if (!!b.favorite !== !!a.favorite) {
        return a.favorite ? -1 : 1;
      }
      return String(a.name).localeCompare(String(b.name), 'zh-Hans-CN');
    });
  }

  function getVisibleContacts() {
    var keyword = state.keyword.trim().toLowerCase();
    var visible = state.contacts.filter(function (item) {
      if (state.onlyFavorite && !item.favorite) {
        return false;
      }
      if (!keyword) {
        return true;
      }
      return [item.name, item.phone, item.email, item.social, item.address, item.note]
        .join(' ')
        .toLowerCase()
        .indexOf(keyword) > -1;
    });
    return sortContacts(visible);
  }

  function render() {
    var visible = getVisibleContacts();
    els.list.innerHTML = '';

    visible.forEach(function (item) {
      els.list.appendChild(createCard(item));
    });

    els.empty.hidden = visible.length > 0;
    els.count.textContent = '共 ' + state.contacts.length + ' 位联系人 · 收藏 ' + countFavorite() + ' 位';
  }

  function countFavorite() {
    return state.contacts.filter(function (item) {
      return item.favorite;
    }).length;
  }

  function createCard(item) {
    var card = document.createElement('article');
    card.className = 'card' + (item.favorite ? ' is-fav' : '');
    card.dataset.id = item.id;

    var avatar = document.createElement('div');
    avatar.className = 'avatar';
    avatar.textContent = String(item.name || '?').slice(0, 1);

    var info = document.createElement('div');
    info.className = 'info';

    var line1 = document.createElement('div');
    line1.className = 'line-1';

    var name = document.createElement('span');
    name.className = 'name';
    name.textContent = item.name;
    line1.appendChild(name);

    if (item.favorite) {
      var tag = document.createElement('span');
      tag.className = 'tag';
      tag.textContent = '已收藏';
      line1.appendChild(tag);
    }

    var line2 = document.createElement('div');
    line2.className = 'line-2';
    [
      ['手机', item.phone],
      ['邮箱', item.email],
      ['社交', item.social],
      ['住址', item.address],
      ['备注', item.note]
    ].forEach(function (pair) {
      if (pair[1]) {
        var span = document.createElement('span');
        span.textContent = pair[0] + '：' + pair[1];
        line2.appendChild(span);
      }
    });

    info.appendChild(line1);
    info.appendChild(line2);

    var actions = document.createElement('div');
    actions.className = 'actions';

    var starBtn = document.createElement('button');
    starBtn.type = 'button';
    starBtn.className = 'star-btn' + (item.favorite ? ' on' : '');
    starBtn.title = item.favorite ? '取消收藏' : '收藏';
    starBtn.textContent = item.favorite ? '★' : '☆';
    starBtn.addEventListener('click', function () {
      toggleFavorite(item.id);
    });

    var editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'link-btn';
    editBtn.textContent = '编辑';
    editBtn.addEventListener('click', function () {
      openDialog(item.id);
    });

    var delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'link-btn';
    delBtn.textContent = '删除';
    delBtn.addEventListener('click', function () {
      removeContact(item.id);
    });

    actions.appendChild(starBtn);
    actions.appendChild(editBtn);
    actions.appendChild(delBtn);

    card.appendChild(avatar);
    card.appendChild(info);
    card.appendChild(actions);
    return card;
  }

  /* ================= 增删改 ================= */

  function toggleFavorite(id) {
    var target = findContact(id);
    if (!target) {
      return;
    }
    target.favorite = !target.favorite;
    saveContacts();
    render();
    toast(target.favorite ? '已加入收藏' : '已取消收藏');
  }

  function removeContact(id) {
    var target = findContact(id);
    if (!target) {
      return;
    }
    if (!window.confirm('确定删除联系人「' + target.name + '」吗？')) {
      return;
    }
    state.contacts = state.contacts.filter(function (item) {
      return item.id !== id;
    });
    saveContacts();
    render();
    toast('已删除');
  }

  function findContact(id) {
    return state.contacts.filter(function (item) {
      return item.id === id;
    })[0];
  }

  function openDialog(id) {
    state.editingId = id || null;
    var data = id ? findContact(id) : null;

    els.dialogTitle.textContent = data ? '编辑联系人' : '新增联系人';
    els.form.name.value = data ? data.name : '';
    els.form.phone.value = data ? data.phone || '' : '';
    els.form.email.value = data ? data.email || '' : '';
    els.form.social.value = data ? data.social || '' : '';
    els.form.address.value = data ? data.address || '' : '';
    els.form.note.value = data ? data.note || '' : '';
    els.form.favorite.checked = data ? !!data.favorite : false;

    els.mask.hidden = false;
    els.form.name.focus();
  }

  function closeDialog() {
    els.mask.hidden = true;
    state.editingId = null;
  }

  function submitForm(event) {
    event.preventDefault();
    var payload = {
      name: els.form.name.value.trim(),
      phone: els.form.phone.value.trim(),
      email: els.form.email.value.trim(),
      social: els.form.social.value.trim(),
      address: els.form.address.value.trim(),
      note: els.form.note.value.trim(),
      favorite: els.form.favorite.checked
    };

    if (!payload.name) {
      toast('请填写姓名');
      return;
    }

    if (state.editingId) {
      var target = findContact(state.editingId);
      if (target) {
        Object.assign(target, payload);
      }
      toast('已保存修改');
    } else {
      payload.id = createId();
      state.contacts.push(payload);
      toast('已新增联系人');
    }

    saveContacts();
    render();
    closeDialog();
  }

  /* ================= 导入导出 ================= */

  function exportExcel() {
    if (!state.contacts.length) {
      toast('暂无可导出的联系人');
      return;
    }
    var rows = sortContacts(state.contacts).map(function (item) {
      return {
        姓名: item.name,
        是否收藏: item.favorite ? '是' : '否',
        手机号: item.phone || '',
        邮箱: item.email || '',
        社交账号: item.social || '',
        住址: item.address || '',
        备注: item.note || ''
      };
    });

    var sheet = XLSX.utils.json_to_sheet(rows, { header: EXCEL_HEADERS });
    sheet['!cols'] = [{ wch: 10 }, { wch: 9 }, { wch: 16 }, { wch: 26 }, { wch: 18 }, { wch: 30 }, { wch: 16 }];
    var book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, '通讯录');
    XLSX.writeFile(book, '通讯录_' + formatDate(new Date()) + '.xlsx');
    toast('已导出 ' + rows.length + ' 位联系人');
  }

  function formatDate(date) {
    var pad = function (num) {
      return num < 10 ? '0' + num : String(num);
    };
    return date.getFullYear() + pad(date.getMonth() + 1) + pad(date.getDate());
  }

  function importExcel(file) {
    var reader = new FileReader();

    reader.onload = function (event) {
      try {
        var book = XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
        var sheet = book.Sheets[book.SheetNames[0]];
        var rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        var added = 0;

        rows.forEach(function (row) {
          var name = String(row['姓名'] || row['name'] || '').trim();
          if (!name) {
            return;
          }
          var favValue = String(row['是否收藏'] || row['收藏'] || '').trim();
          state.contacts.push({
            id: createId(),
            name: name,
            favorite: favValue === '是' || favValue === 'Y' || favValue === 'true' || favValue === '1',
            phone: String(row['手机号'] || row['手机'] || '').trim(),
            email: String(row['邮箱'] || '').trim(),
            social: String(row['社交账号'] || row['社交'] || '').trim(),
            address: String(row['住址'] || row['地址'] || '').trim(),
            note: String(row['备注'] || '').trim()
          });
          added += 1;
        });

        saveContacts();
        render();
        toast(added ? '成功导入 ' + added + ' 位联系人' : '未识别到有效数据，请检查表头');
      } catch (err) {
        toast('文件解析失败，请确认是 Excel 格式');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  /* ================= 提示 ================= */

  var toastTimer = null;

  function toast(message) {
    els.toast.textContent = message;
    els.toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      els.toast.hidden = true;
    }, 2000);
  }

  /* ================= 事件绑定 ================= */

  function bindEvents() {
    els.search.addEventListener('input', function (event) {
      state.keyword = event.target.value;
      render();
    });

    els.favFilter.addEventListener('click', function () {
      state.onlyFavorite = !state.onlyFavorite;
      els.favFilter.classList.toggle('active', state.onlyFavorite);
      render();
    });

    els.addBtn.addEventListener('click', function () {
      openDialog(null);
    });

    els.closeBtn.addEventListener('click', closeDialog);
    els.cancelBtn.addEventListener('click', closeDialog);
    els.form.addEventListener('submit', submitForm);

    els.mask.addEventListener('click', function (event) {
      if (event.target === els.mask) {
        closeDialog();
      }
    });

    els.exportBtn.addEventListener('click', exportExcel);

    els.importBtn.addEventListener('click', function () {
      els.fileInput.click();
    });

    els.fileInput.addEventListener('change', function (event) {
      var file = event.target.files && event.target.files[0];
      if (file) {
        importExcel(file);
      }
      event.target.value = '';
    });

    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape' && !els.mask.hidden) {
        closeDialog();
      }
    });
  }

  /* ================= 启动 ================= */

  loadContacts();
  bindEvents();
  render();
})();
