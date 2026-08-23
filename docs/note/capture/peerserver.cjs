/* 撮影用のシグナリングサーバー。
 * 本番は PeerJS の公開サーバー（0.peerjs.com）を使うが、作業環境からは出られない。
 * 手もとに同じものを立てれば、2台つないだ画面を通しで撮れる。
 *
 *   node docs/note/capture/peerserver.cjs
 *
 * 127.0.0.1 に束ねている。:: に束ねると、IPv6 の無い環境で
 * EAFNOSUPPORT で落ちる（実際に踏んだ）。
 */
const { PeerServer } = require('peer');
PeerServer({ port: 9000, path: '/peer', host: '127.0.0.1' }, () => {
  console.log('peerserver up  http://127.0.0.1:9000/peer');
});
