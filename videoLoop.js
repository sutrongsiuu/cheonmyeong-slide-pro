/**
 * 배경 영상 반복 방식 (편집 미리보기 index.html 과 송출 화면 screen.html 이 함께 쓴다)
 *
 *   normal    : 끝나면 처음부터 다시
 *   pingpong  : 끝까지 재생한 뒤 거꾸로 되감듯 재생하고, 처음에 닿으면 다시 앞으로 재생
 *               브라우저는 음수 재생 속도를 지원하지 않으므로, 실제 시간에 맞춰 뒤로 탐색(seek)하며 보여준다.
 *               키프레임 간격이 넓은 영상은 역재생이 다소 끊겨 보일 수 있다.
 *   crossfade : 끝나기 fade초 전에 처음부터 다시 재생하는 영상을 위에 겹쳐 서서히 나타나게 한다.
 *
 * 이어달리기 (normal · crossfade)
 *   같은 영상을 담은 <video> 두 개가 번갈아 달린다. 하나가 보이며 재생하는 동안 다른 하나는 처음 위치에서
 *   첫 화면까지 준비한 채 숨어서 기다린다. 끝에 다다르면 기다리던 쪽을 재생해 바로 위에 띄우고
 *   (페이드 반복이면 fade초 동안 서서히), 다 끝난 쪽은 숨긴 뒤 처음으로 돌려 다음 차례를 기다린다.
 *   보이는 영상을 되감지(탐색) 않으므로, 키프레임 간격이 넓은 무거운 영상도 반복 경계에서 멈칫하지 않는다.
 *   (브라우저 기본 반복은 끝에서 처음으로 되감는 동안 0.1초 안팎 멈추고, 키프레임이 드문 영상일수록 길어진다)
 *
 *   attach 한 <video>(main)가 계속 이 영상을 대표하는 요소다. 두 번째 요소에는 data-loop-twin 속성이 붙어
 *   querySelector('video:not([data-loop-twin])')는 항상 main을 찾는다.
 *   지금 보이며 재생 중인 요소는 VideoLoop.active(main)로 얻는다 (재생·정지·위치 이동은 이 요소에 한다).
 *
 * 사용: VideoLoop.attach(video, { loop, mode, fade, isFollower, seamless })
 *       video 가 화면에 붙은 뒤(appendChild 이후) 호출한다. 화면에서 빠지면 스스로 정리된다.
 *       isFollower() 가 true 이면 스스로 반복 전환을 시작하지 않는다 (편집 미리보기가 송출 영상을 따라갈 때).
 *       seamless:false 면 normal 반복은 브라우저 기본 반복(video.loop)을 쓴다 (디코더를 하나만 쓴다. 편집 미리보기용).
 *       반환값: 앞으로 이 영상을 대표할 요소. 보통 ↔ 페이드 반복끼리 바꾸면 설정만 바뀌고 main 그대로다.
 *               다른 방식으로 다시 attach 할 때 쌍둥이 쪽이 재생 중이었으면 그 요소가 남는다.
 *
 * 화면 전환 도우미: VideoLoop.whenFrame(video, ms) / VideoLoop.whenReady([img·video...], ms)
 *       새로 나타날 사진·영상의 첫 화면이 준비될 때까지 기다린다 (페이드를 부드럽게 시작하기 위해).
 */
(function(){
  const states = new WeakMap()   // main·쌍둥이 → 같은 상태 객체
  const TWIN_AHEAD = 2.5         // 이어달릴 쌍둥이를 끝나기 몇 초 전에 만들어 첫 화면까지 준비해 둘지

  /** 영상을 완전히 내려 디코더를 바로 돌려준다 */
  function unload(v){
    if (!v) return
    try { v.pause(); v.removeAttribute('src'); v.load() } catch (e) {}
    v.remove()
  }

  /**
   * 반복 설정을 떼어 낸다. 지금 보이며 재생 중인 요소 하나만 남기고 다른 하나는 내린다.
   * 남은 요소를 돌려준다 (쌍둥이가 남으면 data-loop-twin 을 떼어 보통 영상으로 만든다).
   */
  function detach(vid){
    const st = states.get(vid)
    if (!st) return vid
    st.dead = true
    if (st.cleanup) st.cleanup()
    if (st.anim) { try { st.anim.cancel() } catch (e) {} ; st.anim = null }
    clearTimeout(st.prerollTimer)
    const keep = st.a || st.main
    const other = keep === st.main ? st.twin : st.main
    states.delete(st.main)
    if (st.twin) states.delete(st.twin)
    unload(other)
    if (keep !== st.main) {
      keep.removeAttribute('data-loop-twin')
      keep.style.pointerEvents = st.mainPointer || ''
    }
    return keep
  }

  function attach(vid, opts){
    if (!vid) return vid
    opts = opts || {}
    const loop = opts.loop !== false
    const mode = ['pingpong', 'crossfade'].includes(opts.mode) ? opts.mode : 'normal'
    const native = !loop || (mode === 'normal' && opts.seamless === false)
    const fade = Math.max(0.2, Math.min(5, parseFloat(opts.fade) || 1))
    const isFollower = typeof opts.isFollower === 'function' ? opts.isFollower : () => false
    // 이어달리기끼리(보통 ↔ 페이드 반복) 바뀌면 처음 위치에서 기다리던 쪽을 버리지 않고 설정만 바꾼다.
    // 버리고 새로 만들면 다음 반복 경계까지 준비가 끝나지 않아 멈칫한다 (예: 같은 영상을 배경 컨트롤 칸으로 송출)
    const old = states.get(vid)
    if (old && !native && mode !== 'pingpong' && old.mode !== 'pingpong') {
      old.mode = mode
      old.fade = fade
      old.isFollower = isFollower
      // 너무 짧아 브라우저 기본 반복으로 돌리던 영상은 새 기준으로 다시 판단한다
      if (old.native) { old.native = false; old.a.loop = false }
      return old.main
    }
    vid = detach(vid)
    if (native) { vid.loop = loop; return vid }
    vid.loop = false
    const st = { mode, main: vid, fade, isFollower, dead: false, cleanup: null }
    states.set(vid, st)
    if (mode === 'pingpong') setupPingPong(vid, st)
    else setupRelay(vid, st)
    return vid
  }

  // ── 역재생 왕복 ──
  function setupPingPong(vid, st){
    st.a = vid
    let reversing = false
    st.isReversing = () => reversing
    let stopReverse = () => {}

    const startReverse = () => {
      if (st.dead || reversing || st.isFollower()) return
      const dur = vid.duration
      if (!dur || !isFinite(dur)) return
      reversing = true
      let seeking = false
      const t0 = performance.now()
      const onSeeked = () => { seeking = false }
      // 역재생 중에 다른 코드가 재생·구간 이동을 하면 역재생을 멈추고 그쪽을 따른다
      const onPlay = () => stopReverse()
      stopReverse = () => {
        reversing = false
        vid.removeEventListener('seeked', onSeeked)
        vid.removeEventListener('play', onPlay)
        stopReverse = () => {}
      }
      vid.addEventListener('seeked', onSeeked)
      vid.addEventListener('play', onPlay)
      vid.pause()

      const tick = () => {
        if (!reversing) return
        if (st.dead || !vid.isConnected || st.isFollower()) { stopReverse(); if (!vid.isConnected) detach(vid); return }
        const target = dur - (performance.now() - t0) / 1000
        if (target <= 0.04) {
          stopReverse()
          vid.currentTime = 0
          vid.play().catch(() => {})
          return
        }
        // 이전 탐색이 끝난 뒤에만 다음 위치로 이동 (느린 영상은 프레임을 건너뛰며 실제 속도를 유지)
        if (!seeking) { seeking = true; vid.currentTime = target }
        requestAnimationFrame(tick)
      }
      requestAnimationFrame(tick)
    }

    vid.addEventListener('ended', startReverse)
    st.cleanup = () => { vid.removeEventListener('ended', startReverse); stopReverse() }
  }

  /** 첫 화면이 실제로 그려지면 cb (프레임 알림이 없으면 playing 이벤트, 그래도 없으면 ms 뒤) */
  function onFirstFrame(v, cb, ms){
    let done = false
    const fire = () => {
      if (done) return
      done = true
      clearTimeout(timer)
      v.removeEventListener('playing', onPlaying)
      cb()
    }
    const onPlaying = () => requestAnimationFrame(() => requestAnimationFrame(fire))
    const timer = setTimeout(fire, ms || 500)
    if (v.requestVideoFrameCallback) v.requestVideoFrameCallback(() => fire())
    else v.addEventListener('playing', onPlaying)
  }

  // ── 이어달리기 (normal · crossfade) ──
  function setupRelay(vid, st){
    st.a = vid            // 지금 보이며 재생하는 요소
    st.b = null           // 처음 위치에서 기다리는 요소 (쌍둥이는 재생이 시작된 뒤 조금 있다 만든다)
    st.twin = null
    st.phase = 'idle'     // idle | starting(기다리던 쪽 재생 시작) | fading(위에 겹쳐 나타나는 중)
    st.mirroring = false  // 편집 미리보기가 송출 화면의 페이드 반복을 그대로 따라 하는 중
    // 끝나기 몇 초 전에 다음 주자를 출발시킬지
    st.lead = () => st.mode === 'crossfade' ? st.fade : st.startLag
    st.anim = null
    st.fadeStart = 0
    st.startLag = 0.12    // 재생을 시작해 첫 화면이 나올 때까지 걸리는 시간(초). 실제 값으로 계속 보정한다
    // 이보다 짧은 영상은 브라우저 기본 반복으로 (반복 방식이 바뀔 수 있으므로 그때그때 계산)
    st.minDur = () => st.mode === 'crossfade' ? st.fade * 2 + 0.3 : 1
    st.native = false
    // 쌍둥이는 끝나기 조금 전에 만든다: 디코더를 하나 더 여는 일이 화면 전환과 겹치지 않게 재생 시작 후 최소 2.5초 뒤,
    // 그리고 긴 영상이면 대부분의 시간 동안 디코더를 하나만 쓰도록 끝나기 TWIN_AHEAD초 전에
    st.twinAt = performance.now() + 2500
    st.gone = 0

    // 두 요소의 위치·크기는 같게, 투명도와 전환은 이어달리기가 직접 맡는다
    const copyStyle = (from, to, opacity) => {
      to.style.cssText = from.style.cssText
      to.style.transition = 'none'
      to.style.opacity = opacity
      if (to === st.twin) to.style.pointerEvents = 'none'
      else to.style.pointerEvents = st.mainPointer || ''
    }
    st.mainPointer = vid.style.pointerEvents
    // b를 a 바로 위(다음 형제)에 둔다. 다른 영상(예: 그룹 배경이 바뀌며 나타나는 새 영상)보다 위로 올라가지 않게 바로 뒤에만 둔다
    const place = (b, a) => {
      if (a.parentNode && (b.parentNode !== a.parentNode || b.previousSibling !== a)) a.parentNode.insertBefore(b, a.nextSibling)
    }

    const ensureTwin = () => {
      if (st.twin) return st.twin
      const src = st.main.currentSrc || st.main.getAttribute('src') || st.main.src
      if (!src || !st.a.parentNode) return null
      const twin = document.createElement('video')
      twin.muted = true
      twin.preload = 'auto'
      twin.playsInline = true
      twin.setAttribute('data-loop-twin', '1')
      st.twin = twin
      copyStyle(st.a, twin, '0')
      twin.src = src
      st.a.parentNode.insertBefore(twin, st.a.nextSibling)
      listen(twin)
      states.set(twin, st)
      st.b = twin
      return twin
    }

    // 다 달린 쪽은 숨긴 채 처음으로 돌려 다음 차례를 기다린다 (전환이 끝난 뒤 잠시 있다가)
    const schedulePreroll = () => {
      clearTimeout(st.prerollTimer)
      st.prerollTimer = setTimeout(() => {
        const s = st.b
        if (st.dead || st.phase !== 'idle' || !s || !s.paused) return
        if (s.currentTime > 0.01) s.currentTime = 0
      }, 400)
    }

    // 기다리던 쪽이 완전히 보이면 역할을 바꾼다
    const complete = () => {
      if (st.dead) return
      const a = st.a, b = st.b
      if (st.anim) { try { st.anim.cancel() } catch (e) {} ; st.anim = null }
      b.style.opacity = '1'
      // 소리 설정을 넘겨받는다 (소리를 켠 영상이면 이제 새로 달리는 쪽에서 소리가 난다)
      b.muted = a.muted
      a.muted = true
      try { a.pause() } catch (e) {}
      a.style.opacity = '0'
      st.a = b
      st.b = a
      st.phase = 'idle'
      schedulePreroll()
    }

    const cancelRelay = () => {
      if (st.anim) { try { st.anim.cancel() } catch (e) {} ; st.anim = null }
      const b = st.b
      if (b) { try { b.pause() } catch (e) {} ; b.style.opacity = '0' }
      st.phase = 'idle'
      st.mirroring = false
      schedulePreroll()
    }

    /** 기다리던 쪽을 재생해 위에 띄운다. 준비가 안 됐으면 false */
    const beginRelay = () => {
      ensureTwin()
      const a = st.a, b = st.b
      if (!b || b.error || b.readyState < 2) return false
      st.phase = 'starting'
      copyStyle(a, b, '0')
      place(b, a)
      if (b.currentTime > 0.05) b.currentTime = 0
      b.muted = true
      b.playbackRate = 1
      const t0 = performance.now()
      b.play().catch(() => {})
      onFirstFrame(b, () => {
        if (st.dead || st.phase !== 'starting' || st.b !== b) return
        const lag = (performance.now() - t0) / 1000
        st.startLag = Math.max(0.03, Math.min(0.4, st.startLag * 0.6 + lag * 0.4))
        if (st.mode === 'crossfade') {
          st.phase = 'fading'
          st.fadeStart = performance.now()
          // 합성 스레드에서 도는 애니메이션이라 영상 디코딩으로 바빠도 고르게 나타난다
          st.anim = b.animate([{ opacity: 0 }, { opacity: 1 }], { duration: st.fade * 1000, easing: 'linear', fill: 'forwards' })
          st.anim.onfinish = () => { if (st.phase === 'fading' && !st.mirroring) complete() }
        } else {
          complete()
        }
      }, 600)
      return true
    }

    /**
     * 송출 화면의 페이드 반복을 그대로 따라 한다 (편집 미리보기 전용).
     * info = { twinT: 송출에서 새로 달리는 쪽의 지금 위치, p: 진행도 0~1 } | null(송출 쪽 전환 끝남)
     */
    st.mirror = (info, lag) => {
      if (info) {
        if (st.mirroring || st.phase !== 'idle') return
        ensureTwin()
        const a = st.a, b = st.b
        if (!b || b.readyState < 1) return
        st.mirroring = true
        st.phase = 'fading'
        const p = Math.max(0, Math.min(1, +info.p || 0))
        copyStyle(a, b, String(p))
        place(b, a)
        // 처음 위치에 준비된 쪽을 조금만 앞으로 옮기므로 가볍다
        const t = Math.max(0, (+info.twinT || 0) + (lag || 0))
        if (Math.abs(b.currentTime - t) > 0.08) b.currentTime = t
        b.muted = true
        b.playbackRate = 1
        b.play().catch(() => {})
        st.fadeStart = performance.now() - p * st.fade * 1000
        st.anim = b.animate([{ opacity: p }, { opacity: 1 }], { duration: Math.max(50, (1 - p) * st.fade * 1000), easing: 'linear', fill: 'forwards' })
      } else if (st.mirroring) {
        st.mirroring = false
        complete()
      }
    }
    st.isBusy = () => st.phase !== 'idle' || st.mirroring
    st.isMirroring = () => st.mirroring
    st.fadeInfo = () => (st.phase === 'fading' && !st.mirroring && st.b)
      ? { t: st.b.currentTime, p: Math.min(1, (performance.now() - st.fadeStart) / (st.fade * 1000)) }
      : null

    // 보이는 쪽을 다른 코드가 멈추거나 옮기면 진행 중인 전환을 거둔다
    const onPause = e => { if (e.target === st.a && !st.a.ended && st.phase !== 'idle') cancelRelay() }
    const onSeeking = e => { if (e.target === st.a && st.phase !== 'idle') cancelRelay() }
    const onEnded = e => {
      if (e.target !== st.a || st.dead) return
      // 송출 쪽 신호를 놓쳤어도 끝에 닿으면 넘겨받는다
      if (st.mirroring) { st.mirroring = false; complete(); return }
      if (st.phase !== 'idle') return
      if (st.native) { st.a.currentTime = 0; st.a.play().catch(() => {}); return }
      // 미리 시작하지 못했으면 지금이라도 넘겨주고, 기다리는 쪽이 없으면 기본 방식으로 처음부터
      if (!beginRelay()) { st.a.currentTime = 0; st.a.play().catch(() => {}) }
    }
    const listen = el => {
      el.addEventListener('pause', onPause)
      el.addEventListener('seeking', onSeeking)
      el.addEventListener('ended', onEnded)
    }
    listen(vid)
    st.cleanup = () => {
      ;[st.main, st.twin].forEach(el => {
        if (!el) return
        el.removeEventListener('pause', onPause)
        el.removeEventListener('seeking', onSeeking)
        el.removeEventListener('ended', onEnded)
      })
    }

    const tick = () => {
      if (st.dead) return
      // 화면에서 빠졌으면 (두 프레임 연속) 둘 다 내려 디코더를 돌려준다
      if (!st.main.isConnected && !(st.twin && st.twin.isConnected)) {
        if (++st.gone >= 2) { unload(detach(st.main)); return }
      } else st.gone = 0
      const a = st.a, dur = a.duration
      const known = dur && isFinite(dur)
      if (known && !st.native && dur < st.minDur() && st.a === st.main) {
        st.native = true
        a.loop = true
      }
      if (!st.native && st.phase === 'idle' && !st.mirroring && !a.paused) {
        if (!st.twin && known && performance.now() >= st.twinAt && a.currentTime >= dur - st.lead() - TWIN_AHEAD) ensureTwin()
        // 남은 시간이 '시작에 걸리는 시간'(페이드 반복이면 페이드 시간)보다 짧아지면 다음 주자를 출발시킨다
        if (known && !a.seeking && !st.isFollower() && a.currentTime >= dur - st.lead()) beginRelay()
      }
      requestAnimationFrame(tick)
    }
    requestAnimationFrame(tick)
  }

  /** 지금 보이며 재생 중인 요소 (반복 설정이 없으면 그 요소 그대로) */
  function active(vid){ const st = vid && states.get(vid); return (st && st.a) || vid }
  /** 이 영상을 이루는 요소들 (main과 쌍둥이) */
  function elements(vid){ const st = vid && states.get(vid); return st ? [st.main, st.twin].filter(Boolean) : (vid ? [vid] : []) }
  /** 반복 재생하는 영상인지 (기본 반복이든 이어달리기든) */
  function looping(vid){ if (!vid) return false; if (vid.loop) return true; const st = states.get(vid); return !!(st && st.mode !== 'pingpong') }
  /** 역재생 왕복에서 지금 거꾸로 재생 중인지 */
  function reversing(vid){ const st = states.get(vid); return !!(st && st.isReversing && st.isReversing()) }
  /** 반복 전환(이어달리기) 중인지 */
  function busy(vid){ const st = states.get(vid); return !!(st && st.isBusy && st.isBusy()) }

  /** 송출 화면 보고용: 페이드 반복으로 겹치는 중이면 { t: 새로 달리는 쪽 위치, p: 진행도 } */
  function fadeInfo(vid){ const st = states.get(vid); return (st && st.fadeInfo) ? st.fadeInfo() : null }
  /** 편집 미리보기용: 송출 화면의 페이드를 그대로 따라 한다 */
  function mirror(vid, info, lag){ const st = states.get(vid); if (st && st.mirror) st.mirror(info, lag) }
  function mirroring(vid){ const st = states.get(vid); return !!(st && st.isMirroring && st.isMirroring()) }

  /**
   * 영상의 첫 화면이 그려질 준비가 될 때까지 기다린다 (화면 전환 페이드를 시작하기 전에 쓴다).
   * 준비 전에 페이드를 시작하면 빈 화면이 비치다가 영상이 툭 나타나 끊겨 보인다.
   *   - 재생할 영상: 이미 재생 중이면 바로, 아니면 재생이 시작되거나 다음 프레임이 올라올 때
   *   - 멈춰 둘 영상(클릭 시 재생): 첫 프레임이 디코딩되면
   *   - 어느 쪽이든 ms(기본 700ms)가 지나면 더 기다리지 않는다
   */
  function whenFrame(vid, ms){
    vid = active(vid)
    return new Promise(resolve => {
      if (!vid) { resolve(); return }
      let done = false
      const finish = () => {
        if (done) return
        done = true
        clearTimeout(timer)
        vid.removeEventListener('loadeddata', finish)
        vid.removeEventListener('playing', onPlaying)
        resolve()
      }
      const onPlaying = () => requestAnimationFrame(finish)
      const timer = setTimeout(finish, ms || 700)
      if (vid.paused && !vid.autoplay) {
        if (vid.readyState >= 2) finish()
        else vid.addEventListener('loadeddata', finish)
        return
      }
      if (!vid.paused && vid.readyState >= 3 && vid.currentTime > 0) { finish(); return }
      vid.addEventListener('playing', onPlaying)
      if (vid.requestVideoFrameCallback) vid.requestVideoFrameCallback(() => finish())
    })
  }

  /** 여러 사진(<img>)·영상의 첫 화면이 모두 준비될 때까지 기다린다 (최대 ms) */
  function whenReady(list, ms){
    ms = ms || 700
    const tasks = (list || []).filter(Boolean).map(el =>
      el.tagName === 'IMG'
        ? (el.decode ? el.decode().catch(() => {}) : Promise.resolve())
        : whenFrame(el, ms))
    return Promise.race([Promise.all(tasks), new Promise(r => setTimeout(r, ms))])
  }

  window.VideoLoop = { attach, detach, unload, active, elements, looping, reversing, busy, fadeInfo, mirror, mirroring, whenFrame, whenReady }
})()
