import React from "react";
/**
 * This component is used to display the required
 * terms of use statement which can be found in a
 * link in the about tab.
 */
class TermsOfUse extends React.Component {
  render() {
    return (
      <div style={{padding:20}}>
        <h1><img src="/icon.png" width="32"/> AddrTree:階層型アドレス帳のような何か について</h1>
        <p><u>本アプリは現状のまま提供され、不具合その他動作の保証は一切いたしません。</u></p>
        <p><li>階層型アドレス帳のような何かです。</li>
        <li>Teams や Outlook の 「アプリ」 （タブ）として動作します（ 「アドイン」ではありません）。</li>
        <li>Entra ID の「ジョブ情報」 の 「部署」 （department）の情報をもとに階層を表示します。</li>
        <li>管理者はこのアプリに対して Graph の User.Read.All の ユーザに委任された権限を承認する必要があります。</li></p>
        <h2>展開方法</h2>
        <p><ol><li>Microsoft 365 管理センター https://admin.microsoft.com/ を開く</li>
        <li>「設定」の「統合アプリ」の「カスタムアプリをアップロード」をクリック</li>
        <li>アプリの種類「Teamsアプリ」で app.zip をローカルに保存したものをアップロードして「次へ」</li>
        <li>展開するアプリで「階層型アドレス帳のような何か」が表示されたら「次へ」</li>
        <li>ユーザを追加で「特定のユーザまたはグループ」あるいは「組織全体」などをお好みで指定して「次へ」</li>
        <li>「アクセス許可を承認する」をクリック</li>
        <li>ポップアップで管理者としてログインし、User.Read.Allのユーザ委任アクセスを「承諾」して「次へ」</li>
        <li>「展開の完了」をクリック</li></ol></p>
        <p>by Keigo YAMAZAKI (@ymzkei5)</p>
      </div>
    );
  }
}

export default TermsOfUse;
