/**
 * 배경 영상 반복 방식 (편집 미리보기 index.html 과 송출 화면 screen.html 이 함께 쓴다)
 *
 *   normal    : 끝나면 처음부터 다시 (기본 video.loop)
 *   pingpong  : 끝까지 재생한 뒤 거꾸로 되감듯 재생하고, 처음에 닿으면 다시 앞으로 재생
 *               브라우저는 음수 재생 속도를 지원하지 않으므로, 실제 시간에 맞춰 뒤로 탐색(seek)하며 보여준다.
 *               키프레임 간격이 넓은 영상은 역재생이 다소 끊겨 보일 수 있다.
 *   crossfade : 끝나기 fade초 전에 같은 영상의 '쌍둥이'를 처음부터 재생하며 위에 겹쳐 서서히 나타나게 한 뒤,
 *               원래 영상을 쌍둥이 위치로 옮기고 쌍둥이를 숨긴다. 원래 영상 요소가 항상 기준이므로
 *               querySelector('video') 로 영상을 찾는 기존 코드(진행 바·싱크·클릭 재생)는 그대로 동작한다.
 *
 * 사용: VideoLoop.attach(video, { loop, mode, fade, isFollower })
 *       video 가 화면에 붙은 뒤(appendChild 이후) 호출한다. 화면에서 빠지면 스스로 정리된다.
 *       isFollower() 가 true 이면 역재생을 직접 하지 않는다 (편집 미리보기가 송출 영상을 따라갈 때).
 */
(function(){
  const states = new WeakMap()

  function detach(vid){
    const st = states.get(vid)
    if (!st) return
    st.dead = true
    if (st.cleanup) st.cleanup()
    if (st.twin) { try { st.twin.pause() } catch (e) {} ; st.twin.remove() }
    states.delete(vid)
  }

  function attach(vid, opts){
    if (!vid) return
    detach(vid)
    opts = opts || {}
    const loop = opts.loop !== false
    const mode = ['pingpong', 'crossfade'].includes(opts.mode) ? opts.mode : 'normal'
    if (!loop || mode === 'normal') { vid.loop = loop; return }
    vid.loop = false
    const st = {
      mode,
      fade: Math.max(0.2, Math.min(5, parseFloat(opts.fade) || 1)),
      isFollower: typeof opts.isFollower === 'function' ? opts.isFollower : () => false,
      dead: false,
      twin: null,
      cleanup: null
    }
    states.set(vid, st)
    if (mode === 'pingpong') setupPingPong(vid, st)
    else setupCrossfade(vid, st)
  }

  // ── 역재생 왕복 ──
  function setupPingPong(vid, st){
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

  // ── 페이드 반복 ──
  function setupCrossfade(vid, st){
    const twin = document.createElement('video')
    twin.muted = true
    twin.preload = 'auto'
    twin.playsInline = true
    twin.setAttribute('data-loop-twin', '1')
    twin.src = vid.currentSrc || vid.getAttribute('src') || vid.src
    st.twin = twin

    let fading = false, handoff = false, fadeStart = 0, fadeP = 0
    // 따라가기(미러) 상태: 편집 미리보기가 송출 화면의 페이드를 그대로 따라 할 때
    //   false | true(겹치는 중) | 'ending'(원래 영상으로 넘겨받는 중)
    let mirroring = false, mirrorStart = 0
    st.isBusy = () => fading || handoff || mirroring === 'ending'
    st.isMirroring = () => !!mirroring
    st.fadeInfo = () => (fading && !handoff) ? { t: twin.currentTime, p: fadeP } : null

    // 원래 영상을 쌍둥이 위치로 옮기고, 실제 프레임이 그려지면 쌍둥이를 숨긴다
    const handBack = (onDone) => {
      vid.currentTime = Math.min(twin.currentTime + 0.03, (vid.duration || 1) - 0.1)
      vid.play().catch(() => {})
      let finished = false
      const done = () => {
        if (finished) return
        finished = true
        twin.style.opacity = '0'
        try { twin.pause() } catch (e) {}
        onDone()
      }
      if (vid.requestVideoFrameCallback) vid.requestVideoFrameCallback(() => done())
      else vid.addEventListener('playing', done, { once: true })
      setTimeout(done, 600) // 프레임 콜백이 오지 않는 경우 대비
    }

    /**
     * 송출 화면의 페이드 상태를 그대로 따라 한다 (편집 미리보기 전용).
     * info = { twinT: 송출 쌍둥이 영상의 지금 위치, p: 진행도 0~1 } | null(송출 페이드 끝남)
     */
    st.mirror = (info, lag) => {
      if (fading) return
      if (info) {
        if (!mirroring || mirroring === 'ending') {
          mirroring = true
          place()
          twin.style.opacity = String(info.p)
          twin.currentTime = Math.max(0, info.twinT + (lag || 0))
          twin.play().catch(() => {})
          mirrorStart = performance.now() - info.p * st.fade * 1000
        } else if (Math.abs(twin.currentTime - info.twinT) > 0.25) {
          twin.currentTime = Math.max(0, info.twinT + (lag || 0))
        }
      } else if (mirroring === true) {
        mirroring = 'ending'
        handBack(() => { if (mirroring === 'ending') mirroring = false })
      }
    }

    const place = () => {
      if (vid.parentNode && (twin.parentNode !== vid.parentNode || twin.previousSibling !== vid)) {
        vid.parentNode.insertBefore(twin, vid.nextSibling)
      }
      twin.style.cssText = vid.style.cssText + ';pointer-events:none;transition:none;opacity:' + (twin.style.opacity || 0) + ';'
    }
    const cancelFade = () => {
      if (!fading) return
      fading = false; handoff = false
      twin.style.opacity = '0'
      try { twin.pause() } catch (e) {}
    }
    // 사용자가 구간을 옮기거나 멈추면 진행 중인 페이드를 취소
    const onSeeking = () => { if (!handoff && mirroring !== 'ending') cancelFade() }
    const onPause = () => { if (!vid.ended && !handoff) cancelFade() }
    vid.addEventListener('seeking', onSeeking)
    vid.addEventListener('pause', onPause)
    st.cleanup = () => { vid.removeEventListener('seeking', onSeeking); vid.removeEventListener('pause', onPause) }

    const tick = () => {
      if (st.dead) return
      if (!vid.isConnected) { detach(vid); return }
      const dur = vid.duration
      if (mirroring === true) {
        twin.style.opacity = String(Math.min(1, (performance.now() - mirrorStart) / (st.fade * 1000)))
      } else if (!fading) {
        // 송출 영상을 따라가는 편집 미리보기는 스스로 페이드를 시작하지 않고 송출 쪽 페이드를 따라 한다
        if (!mirroring && !st.isFollower() && dur && isFinite(dur) && dur > st.fade * 2 && !vid.paused && vid.currentTime >= dur - st.fade) {
          fading = true
          fadeStart = performance.now()
          twin.style.opacity = '0'
          place()
          twin.currentTime = 0
          twin.play().catch(() => {})
        }
      } else {
        const p = Math.min(1, (performance.now() - fadeStart) / (st.fade * 1000))
        fadeP = p
        twin.style.opacity = String(p)
        if (p >= 1 && !handoff) {
          handoff = true
          handBack(() => { handoff = false; fading = false })
        }
      }
      requestAnimationFrame(tick)
    }
    requestAnimationFrame(tick)
  }

  /** 역재생 왕복에서 지금 거꾸로 재생 중인지 */
  function reversing(vid){ const st = states.get(vid); return !!(st && st.isReversing && st.isReversing()) }
  /** 페이드 반복으로 처음과 끝을 겹치는 중인지 */
  function busy(vid){ const st = states.get(vid); return !!(st && st.isBusy && st.isBusy()) }

  /** 송출 화면 보고용: 페이드 반복으로 겹치는 중이면 { t: 쌍둥이 위치, p: 진행도 } */
  function fadeInfo(vid){ const st = states.get(vid); return (st && st.fadeInfo) ? st.fadeInfo() : null }
  /** 편집 미리보기용: 송출 화면의 페이드를 그대로 따라 한다 */
  function mirror(vid, info, lag){ const st = states.get(vid); if (st && st.mirror) st.mirror(info, lag) }
  function mirroring(vid){ const st = states.get(vid); return !!(st && st.isMirroring && st.isMirroring()) }

  window.VideoLoop = { attach, detach, reversing, busy, fadeInfo, mirror, mirroring }
})()
