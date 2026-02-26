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

  // Collaboration: Presence, Comments, Activities
  var presenceRef = null;
  var commentsRef = null;
  var activitiesRef = null;
  var unsubscribePresence = null;
  var unsubscribeComments = null;
  var unsubscribeActivities = null;
  var presenceInterval = null;
  var onPresenceUpdatedCallback = null;
  var onCommentsUpdatedCallback = null;
  var onActivitiesUpdatedCallback = null;
  var cachedPresenceUsers = [];

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
        initCollaborationRefs();
        subscribeToPresence();
        subscribeToActivities();
      } else {
        detachStoreListener();
        unsubscribeFromPresence();
        removePresence();
        if (unsubscribeComments) {
          unsubscribeComments();
          unsubscribeComments = null;
        }
        if (unsubscribeActivities) {
          unsubscribeActivities();
          unsubscribeActivities = null;
        }
        cachedPresenceUsers = [];
        if (typeof window.clearFirestoreStore === 'function') window.clearFirestoreStore();
        if (typeof window.clearCollaborationState === 'function') window.clearCollaborationState();
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

  // ── Presence System ─────────────────────────────────────────────────────────

  function initCollaborationRefs() {
    if (!db) return;
    presenceRef = storeDocRef.collection('presence');
    commentsRef = storeDocRef.collection('comments');
    activitiesRef = storeDocRef.collection('activities');
  }

  function stringToColorFirebase(str) {
    if (!str) return 'hsl(0, 0%, 50%)';
    var hash = 0;
    for (var i = 0; i < str.length; i++) {
      hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    var hue = Math.abs(hash) % 360;
    return 'hsl(' + hue + ', 70%, 50%)';
  }

  function updatePresence(activeWeekId) {
    if (!presenceRef || !auth || !auth.currentUser) return;
    var user = auth.currentUser;
    var presenceData = {
      odingUserId: user.uid,
      displayName: user.displayName || user.email || 'Anonymous',
      email: user.email || '',
      photoURL: user.photoURL || '',
      color: stringToColorFirebase(user.displayName || user.email),
      lastSeen: firebase.firestore.FieldValue.serverTimestamp(),
      activeWeekId: activeWeekId || ''
    };
    presenceRef.doc(user.uid).set(presenceData, { merge: true }).catch(function(err) {
      console.error('Presence update error:', err);
    });
  }

  function removePresence() {
    if (!presenceRef || !auth || !auth.currentUser) return;
    presenceRef.doc(auth.currentUser.uid).delete().catch(function(err) {
      console.error('Presence remove error:', err);
    });
  }

  function subscribeToPresence() {
    if (unsubscribePresence || !presenceRef) return;
    unsubscribePresence = presenceRef.onSnapshot(function(snapshot) {
      var users = [];
      var now = Date.now();
      var twoMinutesAgo = now - (2 * 60 * 1000);
      snapshot.forEach(function(doc) {
        var data = doc.data();
        var lastSeen = data.lastSeen ? data.lastSeen.toMillis() : 0;
        if (lastSeen > twoMinutesAgo) {
          users.push({
            odingUserId: data.odingUserId,
            displayName: data.displayName,
            email: data.email,
            photoURL: data.photoURL,
            color: data.color,
            lastSeen: lastSeen,
            activeWeekId: data.activeWeekId
          });
        }
      });
      cachedPresenceUsers = users;
      if (typeof onPresenceUpdatedCallback === 'function') {
        onPresenceUpdatedCallback(users);
      }
    }, function(err) {
      console.error('Presence snapshot error:', err);
    });
  }

  function unsubscribeFromPresence() {
    if (unsubscribePresence) {
      unsubscribePresence();
      unsubscribePresence = null;
    }
    if (presenceInterval) {
      clearInterval(presenceInterval);
      presenceInterval = null;
    }
  }

  function startPresenceHeartbeat(getActiveWeekId) {
    if (presenceInterval) clearInterval(presenceInterval);
    updatePresence(getActiveWeekId ? getActiveWeekId() : '');
    presenceInterval = setInterval(function() {
      updatePresence(getActiveWeekId ? getActiveWeekId() : '');
    }, 30000);
  }

  function getPresenceUsers() {
    return cachedPresenceUsers;
  }

  // ── Comments System ─────────────────────────────────────────────────────────

  function subscribeToComments(weekId) {
    if (unsubscribeComments) {
      unsubscribeComments();
      unsubscribeComments = null;
    }
    if (!commentsRef || !weekId) return;
    unsubscribeComments = commentsRef
      .where('weekId', '==', weekId)
      .orderBy('createdAt', 'asc')
      .onSnapshot(function(snapshot) {
        var comments = [];
        snapshot.forEach(function(doc) {
          var data = doc.data();
          comments.push({
            id: doc.id,
            weekId: data.weekId,
            itemId: data.itemId,
            itemType: data.itemType,
            authorUid: data.authorUid,
            authorName: data.authorName,
            authorColor: data.authorColor,
            text: data.text,
            createdAt: data.createdAt ? data.createdAt.toMillis() : Date.now(),
            resolved: data.resolved || false
          });
        });
        if (typeof onCommentsUpdatedCallback === 'function') {
          onCommentsUpdatedCallback(comments);
        }
      }, function(err) {
        console.error('Comments snapshot error:', err);
      });
  }

  function addComment(commentData) {
    if (!commentsRef || !auth || !auth.currentUser) return Promise.reject('Not authenticated');
    var user = auth.currentUser;
    var payload = {
      weekId: commentData.weekId,
      itemId: commentData.itemId,
      itemType: commentData.itemType,
      authorUid: user.uid,
      authorName: user.displayName || user.email || 'Anonymous',
      authorColor: stringToColorFirebase(user.displayName || user.email),
      text: commentData.text,
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      resolved: false
    };
    return commentsRef.add(payload);
  }

  function resolveComment(commentId) {
    if (!commentsRef) return Promise.reject('Not initialized');
    return commentsRef.doc(commentId).update({ resolved: true });
  }

  function deleteComment(commentId) {
    if (!commentsRef) return Promise.reject('Not initialized');
    return commentsRef.doc(commentId).delete();
  }

  // ── Activities System ───────────────────────────────────────────────────────

  function subscribeToActivities() {
    if (unsubscribeActivities || !activitiesRef) return;
    var sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    unsubscribeActivities = activitiesRef
      .where('timestamp', '>', sevenDaysAgo)
      .orderBy('timestamp', 'desc')
      .limit(50)
      .onSnapshot(function(snapshot) {
        var activities = [];
        snapshot.forEach(function(doc) {
          var data = doc.data();
          activities.push({
            id: doc.id,
            weekId: data.weekId,
            type: data.type,
            actorUid: data.actorUid,
            actorName: data.actorName,
            actorColor: data.actorColor,
            action: data.action,
            targetText: data.targetText,
            timestamp: data.timestamp ? data.timestamp.toMillis() : Date.now(),
            mentions: data.mentions || []
          });
        });
        if (typeof onActivitiesUpdatedCallback === 'function') {
          onActivitiesUpdatedCallback(activities);
        }
      }, function(err) {
        console.error('Activities snapshot error:', err);
      });
  }

  function logActivity(activityData) {
    if (!activitiesRef || !auth || !auth.currentUser) return Promise.resolve();
    var user = auth.currentUser;
    var payload = {
      weekId: activityData.weekId || '',
      type: activityData.type,
      actorUid: user.uid,
      actorName: user.displayName || user.email || 'Anonymous',
      actorColor: stringToColorFirebase(user.displayName || user.email),
      action: activityData.action,
      targetText: (activityData.targetText || '').substring(0, 50),
      timestamp: firebase.firestore.FieldValue.serverTimestamp(),
      mentions: activityData.mentions || []
    };
    return activitiesRef.add(payload).catch(function(err) {
      console.error('Activity log error:', err);
    });
  }

  // ── Collaboration Callbacks ─────────────────────────────────────────────────

  function setOnPresenceUpdatedCallback(fn) {
    onPresenceUpdatedCallback = fn;
  }

  function setOnCommentsUpdatedCallback(fn) {
    onCommentsUpdatedCallback = fn;
  }

  function setOnActivitiesUpdatedCallback(fn) {
    onActivitiesUpdatedCallback = fn;
  }

  window.firebaseAuth = {
    init: initFirebase,
    signIn: toggleFirebaseAuth,
    signOut: firebaseSignOut,
    getCurrentUser: getCurrentUser,
    isSignedIn: isSignedIn,
    saveStoreToFirestore: saveStoreToFirestore,
    setOnStoreUpdatedCallback: setOnStoreUpdatedCallback,
    // Presence
    updatePresence: updatePresence,
    removePresence: removePresence,
    subscribeToPresence: subscribeToPresence,
    startPresenceHeartbeat: startPresenceHeartbeat,
    getPresenceUsers: getPresenceUsers,
    setOnPresenceUpdatedCallback: setOnPresenceUpdatedCallback,
    // Comments
    subscribeToComments: subscribeToComments,
    addComment: addComment,
    resolveComment: resolveComment,
    deleteComment: deleteComment,
    setOnCommentsUpdatedCallback: setOnCommentsUpdatedCallback,
    // Activities
    subscribeToActivities: subscribeToActivities,
    logActivity: logActivity,
    setOnActivitiesUpdatedCallback: setOnActivitiesUpdatedCallback
  };
  window.toggleFirebaseAuth = toggleFirebaseAuth;
  window.firebaseSignOut = firebaseSignOut;
})();
