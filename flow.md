flowchart TD
    A["1. Identity Provider<br/>Microsoft Entra ID / Okta"] -->|"SCIM / 自動同期"| B["2. Databricks Account Console<br/>アカウントレベルで ID 管理"]

    B --> C["3. Account-level Identities<br/>Users / Groups / Service Principals"]

    C --> D{"権限付与の対象を決める"}

    D -->|"推奨"| E["Groups<br/>ユーザ権限管理の基本単位"]
    D -->|"本番ジョブ / 自動処理"| F["Service Principals"]
    D -->|"非推奨: 例外的に使用"| G["Individual Users"]

    E --> H["4. Workspace Access<br/>ワークスペースへの参加権限"]
    F --> H
    G --> H

    H --> I["5. Workspace<br/>Notebook / Job / Dashboard / Query"]

    I --> J{"管理者ロール"}

    J --> K["Workspace Admin<br/>単一ワークスペースの管理者"]
    J --> L["Account Admin<br/>アカウント全体の管理者"]

    L -->|"作成 / 管理"| I
    L -->|"作成 / 紐付け"| M["Unity Catalog Metastore<br/>リージョン単位のガバナンス境界"]
    L -->|"付与可能"| K
    L -->|"必要に応じて付与"| N["Metastore Admin<br/>任意・強権限"]

    K -->|"管理"| O["Workspace Membership<br/>ユーザ / グループ / SP 割当"]
    K -->|"管理"| P["Workspace Objects<br/>Notebook / Job / Dashboard / Query"]
    K -->|"管理"| Q["Jobs / Run as"]

    I -->|"Unity Catalog 有効化時に接続"| M

    M --> R["6. Catalog"]
    R --> S["7. Schema"]
    S --> T["8. Tables / Views / Volumes / Functions"]

    E -->|"GRANT<br/>USE CATALOG / USE SCHEMA / SELECT / MODIFY"| R
    E -->|"GRANT<br/>USE SCHEMA / SELECT / MODIFY など"| S
    E -->|"必要な場合のみ直接 GRANT"| T

    U["Catalog Owner<br/>推奨: データドメイン管理グループ"] -->|"権限管理 / 所有権管理"| R
    N -->|"広域管理 / 所有権変更"| R
    N -->|"広域管理 / 所有権変更"| S
    N -->|"広域管理 / 所有権変更"| T

    K -. "自動 UC 有効化ワークスペースでは<br/>CREATE CATALOG / CREATE EXTERNAL LOCATION<br/>CREATE STORAGE CREDENTIAL などが付与される場合あり" .-> M

    K -. "Workspace Catalog がある場合<br/>デフォルト所有者" .-> V["Workspace Catalog"]
    V --> S

    T --> W["9. Data Access<br/>実際のデータ参照 / 更新"]
