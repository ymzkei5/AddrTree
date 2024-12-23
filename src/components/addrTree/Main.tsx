import React, { useState, useContext, useEffect, useRef } from 'react';
import { 
  Stack, 
  TextField, 
  DefaultButton, 
  PrimaryButton, 
  IStackTokens, 
  DetailsList, 
  IColumn, 
  Spinner, 
  MessageBar, 
  MessageBarType,
  Selection,
  IObjectWithKey
} from '@fluentui/react';
import { TeamsFxContext } from "../Context";
import { TeamsUserCredential } from '@microsoft/teamsfx';

export type TreeNode = {
  id: string;
  label: string;
  fullLabel: string;
  children?: TreeNode[];
};

export type User = {
  id: string;
  displayName: string;
  department: string;
  mail: string;
  jobTitle: string;
};

const columns: IColumn[] = [
  { key: 'column1', name: '名前', fieldName: 'displayName', minWidth: 100, maxWidth: 150, isResizable: true },
  { key: 'column2', name: '部署', fieldName: 'department', minWidth: 100, maxWidth: 420, isResizable: true },
  { key: 'column3', name: 'メール', fieldName: 'mail', minWidth: 100, maxWidth: 250, isResizable: true },
  { key: 'column4', name: '役職', fieldName: 'jobTitle', minWidth: 100, maxWidth: 150, isResizable: true },
];

const stackTokens: IStackTokens = { childrenGap: 8, padding: 8 };

// キャッシュの有効期限（ミリ秒単位、ここでは1日）
const CACHE_EXPIRATION = 24 * 60 * 60 * 1000; // 1日

// キャッシュキーの定義
const CACHE_KEY_DEPARTMENTS = 'departments_cache';

/**
 * データをキャッシュから取得します。
 * 有効期限をチェックし、データが有効であれば返します。
 * @returns キャッシュされたデータか、null
 */
const getCachedData = (key: string) => {
  const cached = localStorage.getItem(key);
  if (!cached) return null;

  try {
    const parsed = JSON.parse(cached);
    const now = new Date().getTime();
    if (now - parsed.timestamp < CACHE_EXPIRATION) {
      return parsed.data;
    } else {
      // キャッシュが期限切れの場合は削除
      localStorage.removeItem(key);
      return null;
    }
  } catch (error) {
    console.error('キャッシュの解析に失敗しました:', error);
    localStorage.removeItem(key);
    return null;
  }
};

/**
 * データをキャッシュに保存します。
 * @param key キャッシュキー
 * @param data 保存するデータ
 */
const setCachedData = (key: string, data: any) => {
  const cacheEntry = {
    data,
    timestamp: new Date().getTime()
  };
  try {
    localStorage.setItem(key, JSON.stringify(cacheEntry));
  } catch (error) {
    console.error('キャッシュの保存に失敗しました:', error);
  }
};

export const Main: React.FC = () => {
  const [treeData, setTreeData] = useState<TreeNode[]>([]);
  const [expandedNodes, setExpandedNodes] = useState<{ [key: string]: boolean }>({});
  const [selectedNodeId, setSelectedNodeId] = useState<string | null>(null);
  const [selectedNodeLabel, setSelectedNodeLabel] = useState<string | null>(null);
  const [items, setItems] = useState<User[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  
  // 検索入力の状態
  const [searchValue, setSearchValue] = useState<string>('');

  // メールフィールドの状態を文字列に変更
  const [toField, setToField] = useState<string>('');
  const [ccField, setCcField] = useState<string>('');
  const [bccField, setBccField] = useState<string>('');

  // 選択されたアイテムの状態
  const [selectedItems, setSelectedItems] = useState<User[]>([]);

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential! as TeamsUserCredential;

  let idCounter = 1;

  const buildTree = (lines: string[]): TreeNode[] => {
    const root: TreeNode = { id: 'root', label: '組織', fullLabel: '組織', children: [] };
    
    if (lines.length <= 0) {
      return [{id:"1", label:"⚠ 部署 （department）が未設定", fullLabel:"allusers", children:[]}];
    }

    lines.forEach(line => {
      const parts = line.trim().split(' ');
      let currentNode = root;

      parts.forEach(part => {
        if (!currentNode.children) {
          currentNode.children = [];
        }
        let childNode = currentNode.children.find(child => child.label === part);
        if (!childNode) {
          childNode = {
            id: (idCounter++).toString(),
            label: part,
            fullLabel: currentNode.fullLabel === '組織' ? part : `${currentNode.fullLabel} ${part}`,
            children: []
          };
          currentNode.children.push(childNode);
        }
        currentNode = childNode;
      });
    });

    return root.children || [];
  };

  useEffect(() => {
    const fetchTreeData = async () => {
      try {
        // キャッシュからデータを取得
        const cachedDepartments: string[] | null = getCachedData(CACHE_KEY_DEPARTMENTS);
        if (cachedDepartments) {
          console.log('キャッシュから部署データを取得しました。');
          const tree = buildTree(cachedDepartments);
          setTreeData(tree);
          return;
        }

        // キャッシュがない場合はMicrosoft Graph APIから取得
        let accessToken = '';
        try {
          let tokenResponse = await teamsUserCredential.getToken(["User.Read.All"]);
          if (tokenResponse) accessToken = tokenResponse.token;
        } catch (err) {
          console.warn('Token取得中にエラーが発生しました:', err);
        }
        if (accessToken === '') {
          const loginResponse = await teamsUserCredential.login(["User.Read.All"]);
          let tokenResponse = await teamsUserCredential.getToken(["User.Read.All"]);
          if (!tokenResponse) {
            throw new Error('アクセストークンの取得に失敗しました');
          }
          accessToken = tokenResponse.token;
        }

        let departments: string[] = [];
        let url = 'https://graph.microsoft.com/v1.0/users?$select=department&$top=999';

        while (url) {
          const response = await fetch(url, {
            headers: {
              'Authorization': `Bearer ${accessToken}`
            }
          });
          if (!response.ok) {
            throw new Error(`Graph APIの呼び出しに失敗しました: ${response.status} ${response.statusText}`);
          }
          const data = await response.json();
          const users = data.value || [];
          users.forEach((user: any) => {
            if (user.department && user.department.trim() !== '') {
              departments.push(user.department.trim());
            }
          });
          url = data['@odata.nextLink'] || '';
        }

        // 重複を除外しソート
        const uniqueDepartments = Array.from(new Set(departments)).sort();

        // キャッシュに保存
        if (uniqueDepartments.length > 0) {
          setCachedData(CACHE_KEY_DEPARTMENTS, uniqueDepartments);
        }

        const tree = buildTree(uniqueDepartments);
        setTreeData(tree);
      } catch (error: any) {
        console.error(error);
        setErrorMessage('ツリーデータの読み込みに失敗しました');
      }
    };
    fetchTreeData();
  }, [teamsUserCredential]);

  const toggleNode = (nodeId: string) => {
    setExpandedNodes(prev => ({
      ...prev,
      [nodeId]: !prev[nodeId]
    }));
  };

  const onNodeClick = (nodeId: string, label: string) => {
    setSelectedNodeId(nodeId);
    setSelectedNodeLabel(label);
  };

  const renderTree = (nodes: TreeNode[], level: number = 0): React.ReactNode => {
    return (
      <ul style={{ listStyleType: 'none', margin: 0, paddingLeft: level * 10, whiteSpace: 'nowrap' }}>
        {nodes.map((node) => {
          const hasChildren = node.children && node.children.length > 0;
          const isExpanded = expandedNodes[node.id] ?? false;
          const isSelected = selectedNodeId === node.id;

          return (
            <li key={node.id} style={{ cursor: 'pointer' }}>
              {/* 展開/折り畳みボタン */}
              {hasChildren && (
                <span
                  style={{ marginRight: 4 }}
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleNode(node.id);
                  }}
                >
                  {isExpanded ? '[-]' : '[+]'}
                </span>
              )}
              {!hasChildren && <span style={{ marginRight: 4 }}></span>}

              {/* ノードラベル部分 */}
              <span
                onClick={() => onNodeClick(node.id, node.fullLabel)}
                onDoubleClick={(e) => {
                  e.stopPropagation();
                  if (hasChildren) {
                    toggleNode(node.id);
                  }
                }}
                style={{
                  backgroundColor: isSelected ? '#cce5ff' : 'transparent', 
                  display: 'inline-block',
                  padding: '2px 4px',
                  userSelect: 'none'
                }}
                title={node.label} // 長い名前の際にツールチップとして表示可能
              >
                {node.label}
              </span>

              {hasChildren && isExpanded && renderTree(node.children!, level + 1)}
            </li>
          );
        })}
      </ul>
    );
  };

  /**
   * OData クエリで使用する文字列をエスケープします。
   * シングルクォートを2つのシングルクォートに置き換えます。
   * @param value エスケープする文字列
   * @returns エスケープされた文字列
   */
  const escapeODataString = (value: string): string => {
    return value.replace(/'/g, "''");
  };

  /**
   * フィールドをパースして配列に変換します。
   * セミコロンで区切られたメールアドレスを配列に分割します。
   * @param field フィールドの文字列
   * @returns パースされたメールアドレスの配列
   */
  const parseEmailField = (field: string): string[] => {
    return field
      .split(';')
      .map(entry => entry.trim())
      .filter(entry => entry !== '');
  };

  /**
   * ユーザーオブジェクトを "Name <email>" 形式にフォーマットします。
   * @param user ユーザーオブジェクト
   * @returns フォーマットされた文字列
   */
  const formatUser = (user: User): string => {
    return `${user.displayName} <${user.mail}>`;
  };

  /**
   * フィールドにユーザーを追加します。重複を避けます。
   * @param fieldSetter フィールドのセット関数（setToField, setCcField, setBccField）
   * @param fieldValue 現在のフィールドの値
   * @param users 追加するユーザーの配列
   */
  const addUsersToField = (
    fieldSetter: React.Dispatch<React.SetStateAction<string>>,
    fieldValue: string,
    users: User[]
  ) => {
    const existingEntries = parseEmailField(fieldValue);
    const newEntries = users
      .filter(user => user.displayName && user.mail)
      .map(formatUser)
      .filter(formattedUser => !existingEntries.includes(formattedUser));
    
    if (newEntries.length === 0) return;
    
    const updatedEntries = [...existingEntries, ...newEntries];
    fieldSetter(updatedEntries.join('; '));
  };

  const handleAddToTo = () => {
    if (selectedItems.length === 0) {
      alert('ユーザーを選択してください。');
      return;
    }
    addUsersToField(setToField, toField, selectedItems);
  };

  const handleAddToCc = () => {
    if (selectedItems.length === 0) {
      alert('ユーザーを選択してください。');
      return;
    }
    addUsersToField(setCcField, ccField, selectedItems);
  };

  const handleAddToBcc = () => {
    if (selectedItems.length === 0) {
      alert('ユーザーを選択してください。');
      return;
    }
    addUsersToField(setBccField, bccField, selectedItems);
  };

  /**
   * メール作成時の処理
   */
  const handleCreateEmail = () => {
    const toEmails = toField;
    const ccEmails = ccField;
    const bccEmails = bccField;

    const mailto = `mailto:?` +
      (toEmails ? `to=${encodeURIComponent(toEmails)}` : '') +
      (ccEmails ? `&cc=${encodeURIComponent(ccEmails)}` : '') +
      (bccEmails ? `&bcc=${encodeURIComponent(bccEmails)}` : '');
    window.location.href = mailto;
  };

  /**
   * イベント作成時の処理
   */
  const handleCreateEvent = () => {
    const baseUrl = 'https://outlook.office.com/calendar/deeplink/compose';
    const params = new URLSearchParams({
      to: toField,
      cc: ccField,
    });

    const eventUrl = `${baseUrl}?${params.toString()}`;
    window.open(eventUrl, '_blank');
  };

  /**
   * DetailsList のアイテムをダブルクリックしたときに宛先に追加します。
   * Fluent UI の DetailsList では onItemInvoked が使用できます。
   * これはアイテムがダブルクリックまたは Enter キーで「呼び出された」ときにトリガーされます。
   */
  const handleItemInvoked = (item: User) => {
    addUsersToField(setToField, toField, [item]);
  };

  /**
   * テキストフィールドの変更ハンドラ
   * ユーザーが直接フィールドを編集したときにステートを更新します。
   */
  const handleFieldChange = (
    fieldSetter: React.Dispatch<React.SetStateAction<string>>
  ) => (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    fieldSetter(newValue || '');
  };

  /**
   * フィールドのクリア
   */
  const handleClearFields = () => {
    setToField('');
    setCcField('');
    setBccField('');
    // 選択を解除
    selectionRef.current.setAllSelected(false);
    setSelectedItems([]);
  };

  // Selection オブジェクトの useRef の定義
  const selectionRef = useRef<Selection>(new Selection({
    onSelectionChanged: () => {
      const selected = selectionRef.current.getSelection() as User[];
      setSelectedItems(selected);
    }
  }));

  /**
   * fetchUsersByFilter 関数の定義
   */
  const fetchUsersByFilter = async (filterValue: string) => {
    setIsLoading(true);
    setErrorMessage(null);

    try {
      let accessToken = '';
      try {
        let tokenResponse = await teamsUserCredential.getToken(["User.Read.All"]);
        if (tokenResponse) accessToken = tokenResponse.token;
      } catch {}
      if (accessToken === '') {
        await teamsUserCredential.login(["User.Read.All"]);
        let tokenResponse = await teamsUserCredential.getToken(["User.Read.All"]);
        if (!tokenResponse) {
          throw new Error('トークン取得失敗');
        }
        accessToken = tokenResponse.token;
      }

      const filter = (filterValue.includes('allusers') ? '' : `&$filter=${filterValue}`);
      const url = `https://graph.microsoft.com/v1.0/users?$select=id,displayName,department,mail,jobTitle${filter}&$top=999`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        }
      });

      const data = await response.json();
      let allUsers: User[] = data.value || [];

      // ページネーションの処理
      let nextLink = data['@odata.nextLink'];
      while (nextLink) {
        const nextResponse = await fetch(nextLink, {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });
        if (!nextResponse.ok) {
          throw new Error(`Graph APIの呼び出しに失敗しました: ${nextResponse.status} ${nextResponse.statusText}`);
        }
        const nextData = await nextResponse.json();
        allUsers = allUsers.concat(nextData.value || []);
        nextLink = nextData['@odata.nextLink'];
      }

      setItems(allUsers);

      // リストが更新された際に選択を解除
      selectionRef.current.setAllSelected(false);
      setSelectedItems([]);
    } catch (error: any) {
      console.error(error);
      setErrorMessage(error.message || 'ユーザーの取得に失敗しました');
      setItems([]);

      // エラー発生時も選択を解除
      selectionRef.current.setAllSelected(false);
      setSelectedItems([]);
    } finally {
      setIsLoading(false);
    }
  }

  useEffect(() => {
    const fetchUsers = async () => {
      if (!selectedNodeLabel) {
        setItems([]);
        // 選択を解除
        selectionRef.current.setAllSelected(false);
        setSelectedItems([]);
        return;
      }
      const escapedLabel = escapeODataString(selectedNodeLabel);
      const filterValue = encodeURIComponent(`department eq '${escapedLabel}'`);
      fetchUsersByFilter(filterValue);
    };
    fetchUsers();
  }, [selectedNodeLabel, teamsUserCredential]);

  return (
    <Stack 
      verticalFill 
      styles={{ 
        root: { 
          width: '100%', 
          height: '100%', 
          border: '1px solid #ccc', 
          padding: '8px', 
          boxSizing: 'border-box' 
        } 
      }} 
      tokens={stackTokens}
    >
      {/* 上部には検索バーなど */}
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
        <TextField
          label="検索（※前方一致）:"
          underlined
          value={searchValue}
          onChange={(e, newValue) => setSearchValue(newValue || '')}
          onKeyPress={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              // トリガー検索
              const trimmedValue = searchValue.trim();
              if (trimmedValue === '') {
                setItems([]);
                // 選択を解除
                selectionRef.current.setAllSelected(false);
                setSelectedItems([]);
                return;
              }
              const escapedLabel = escapeODataString(trimmedValue);
              const filterValue = encodeURIComponent(`startsWith(department,'${escapedLabel}') OR startsWith(displayName,'${escapedLabel}') OR startsWith(mail,'${escapedLabel}')`);
              fetchUsersByFilter(filterValue);
            }
          }}
        />
        <DefaultButton
          text="検索"
          onClick={() => {
            const trimmedValue = searchValue.trim();
            if (trimmedValue === '') {
              // 空の場合、ユーザーリストをクリア
              setItems([]);
              // 選択を解除
              selectionRef.current.setAllSelected(false);
              setSelectedItems([]);
              return;
            }
            const escapedLabel = escapeODataString(trimmedValue);
            const filterValue = encodeURIComponent(`startsWith(department,'${escapedLabel}') OR startsWith(displayName,'${escapedLabel}') OR startsWith(mail,'${escapedLabel}')`);
            fetchUsersByFilter(filterValue);
          }}
        />
      </Stack>

      <Stack horizontal grow tokens={{ childrenGap: 8 }} styles={{ root: { overflow: 'hidden' } }}>
        {/* 左パネル（ツリー表示） */}
        <Stack styles={{ root: { width: '300px', borderRight: '1px solid #ccc', overflowX: 'auto', overflowY: 'auto' } }}>
          {renderTree(treeData)}
        </Stack>

        {/* 右パネル（テーブル表示） */}
        <Stack grow styles={{ root: { overflowY: 'auto', position: 'relative' } }}>
          {isLoading && (
            <Spinner 
              label="取得中..." 
              styles={{ root: { position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)' } }} 
            />
          )}
          {errorMessage && 
            <MessageBar 
              messageBarType={MessageBarType.error} 
              styles={{ root: { marginBottom: '8px' } }}
            >
              {errorMessage}
            </MessageBar>
          }
          <DetailsList
            items={items}
            columns={columns}
            setKey="set"
            layoutMode={0}
            selection={selectionRef.current} // 正しく current を渡す
            selectionPreservedOnEmptyClick
            isHeaderVisible
            onItemInvoked={handleItemInvoked} // ダブルクリック時のハンドラ
          />
        </Stack>
      </Stack>

      {/* 下部のボタンと入力フィールド */}
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { borderTop: '1px solid #ccc', paddingTop: '8px' } }}>
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          {/* 宛先 */}
          <Stack tokens={{ childrenGap: 4 }} grow>
            <PrimaryButton 
              text="宛先" 
              onClick={handleAddToTo}
              disabled={selectedItems.length === 0}
            />
            <TextField 
              value={toField} 
              multiline 
              onChange={handleFieldChange(setToField)}
            />
          </Stack>

          {/* CC */}
          <Stack tokens={{ childrenGap: 4 }} grow>
            <PrimaryButton 
              text="CC" 
              onClick={handleAddToCc}
              disabled={selectedItems.length === 0}
            />
            <TextField 
              value={ccField} 
              multiline 
              onChange={handleFieldChange(setCcField)}
            />
          </Stack>

          {/* BCC */}
          <Stack tokens={{ childrenGap: 4 }} grow>
            <PrimaryButton 
              text="BCC" 
              onClick={handleAddToBcc}
              disabled={selectedItems.length === 0}
            />
            <TextField 
              value={bccField} 
              multiline 
              onChange={handleFieldChange(setBccField)}
            />
          </Stack>
        </Stack>

        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
          <PrimaryButton 
            text="メール作成" 
            onClick={handleCreateEmail} 
            disabled={toField.trim() === '' && ccField.trim() === '' && bccField.trim() === ''}
          />
          <PrimaryButton 
            text="イベント作成" 
            onClick={handleCreateEvent} 
            disabled={(toField.trim() === '' && ccField.trim() === '') || bccField.trim() !== ''}
          />
          <DefaultButton 
            text="キャンセル" 
            onClick={handleClearFields} 
          />
        </Stack>
      </Stack>
    </Stack>
  );
};