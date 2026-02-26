// ══════════════════════════════════════════════════════════════════════════════
// Firebase: Auth (Google Sign-In) + Firestore (shared store, real-time sync)
// ══════════════════════════════════════════════════════════════════════════════

(function () {
  'use strict';

  var firebaseConfig = {
    apiKey: "AIzaSyB4vIfcojm5fvMIhUMFpxu0nfq1HGn9t60",
    authDomain: "ss-report-900ef.firebaseapp.com",
    projectId: "ss-report-900ef",
    storageBucket: "ss-report-900ef.firebasestorage.app",
    messagingSenderId: "919417174377",
    appId: "1:919417174377:web:d60a60f05f10481e3e65f8"
  };

  var app = null;
  var auth = null;
  var db = null;
  var storeDocRef = null;
  var unsubscribeSnapshot = null;
  var onStoreUpdatedCallback = null;

  function initFirebase() {
    if (app) return;
    app = firebase.initializeApp(firebaseConfig);
    auth = firebase.auth();
    db = firebase.firestore();
    storeDocRef = db.collection('reports').doc('store');

    auth.onAuthStateChanged(function (user) {
      var signInBtn = document.getElementById('googleSignInBtn');
      var authWrap = document.getElementById('authUserWrap');
      var authLabel = document.getElementById('authUserLabel');
      if (signInBtn) signInBtn.style.display = user ? 'none' : '';
      if (authWrap) authWrap.style.display = user ? '' : 'none';
      if (authLabel && user) authLabel.textContent = user.displayName || user.email || 'Signed in';

      if (user) {
        attachStoreListener();
      } else {
        detachStoreListener();
        if (typeof window.clearFirestoreStore === 'function') window.clearFirestoreStore();
      }
    });
  }

  function attachStoreListener() {
    if (unsubscribeSnapshot) return;
    unsubscribeSnapshot = storeDocRef.onSnapshot(
      function (snap) {
        if (!snap.exists()) return;
        var data = snap.data();
        var store = {
          weeks: data.weeks || {},
          order: data.order || [],
          version: data.version || 2
        };
        if (typeof window.setFirestoreStore === 'function') {
          window.setFirestoreStore(store);
          if (typeof onStoreUpdatedCallback === 'function') onStoreUpdatedCallback(store);
        }
      },
      function (err) {
        console.error('Firestore snapshot error:', err);
      }
    );
  }

  function detachStoreListener() {
    if (unsubscribeSnapshot) {
      unsubscribeSnapshot();
      unsubscribeSnapshot = null;
    }
  }

  function toggleFirebaseAuth() {
    if (!auth) initFirebase();
    var provider = new firebase.auth.GoogleAuthProvider();
    auth.signInWithPopup(provider).catch(function (err) {
      console.error('Sign-in error:', err);
      if (typeof window.showToast === 'function') {
        window.showToast(err.message || 'Sign-in failed', 'error');
      }
    });
  }

  function firebaseSignOut() {
    if (auth) auth.signOut();
  }

  function getCurrentUser() {
    return auth ? auth.currentUser : null;
  }

  function isSignedIn() {
    return !!(auth && auth.currentUser);
  }

  function saveStoreToFirestore(storeData, callback) {
    if (!db || !storeDocRef || !auth || !auth.currentUser) {
      if (callback) callback(false);
      return;
    }
    var payload = {
      weeks: storeData.weeks || {},
      order: storeData.order || [],
      version: storeData.version || 2,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: auth.currentUser.uid
    };
    storeDocRef.set(payload, { merge: true }).then(
      function () { if (callback) callback(true); },
      function (err) {
        console.error('Firestore write error:', err);
        if (callback) callback(false);
        if (typeof window.showToast === 'function') window.showToast('Sync failed: ' + (err.message || 'unknown'), 'error');
      }
    );
  }

  function setOnStoreUpdatedCallback(fn) {
    onStoreUpdatedCallback = fn;
  }

  window.firebaseAuth = {
    init: initFirebase,
    signIn: toggleFirebaseAuth,
    signOut: firebaseSignOut,
    getCurrentUser: getCurrentUser,
    isSignedIn: isSignedIn,
    saveStoreToFirestore: saveStoreToFirestore,
    setOnStoreUpdatedCallback: setOnStoreUpdatedCallback
  };
  window.toggleFirebaseAuth = toggleFirebaseAuth;
  window.firebaseSignOut = firebaseSignOut;
})();
