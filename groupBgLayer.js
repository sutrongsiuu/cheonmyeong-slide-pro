/**
 * 그룹 배경 영상 레이어 (편집 미리보기 index.html 과 송출 화면 screen.html 이 함께 쓴다)
 *
 * 그룹 배경이 영상이면, 그 영상을 그대로 쓰는 슬라이드마다 <video>를 새로 만들지 않고
 * A/B 슬라이드 레이어 뒤에 있는 이 레이어의 영상 하나를 계속 이어서 재생한다.
 * 슬라이드를 넘겨도 영상은 처음부터 다시 틀지 않고, 위의 가사·텍스트만 바뀐다.
 *
 *  show(spec, fd) : 이 영상을 보여준다. 이미 같은 영상이면 그대로 이어서 재생한다.
 *                   다른 영상이면 fd초 동안 새 영상이 서서히 나타난다.
 *  hide(fd)       : 따로 배경을 넣은 슬라이드로 넘어갈 때. fd초 동안 사라진 뒤 일시정지한다.
 *                   (다시 show 하면 멈춘 위치부터 이어서 재생)
 *
 * z-index: 평소 0 (두 슬라이드 레이어 아래). 사라지는 동안만 1로 올려, 같은 z-index 1인
 *          '다음 슬라이드 레이어' 위에서 함께 페이드아웃되게 한다. (컨테이너는 LB 뒤에 위치해야 함)
 *
 * spec = { key, url, fit, audio, loop, loopMode, loopFade, autoPlay }
 * opts = { takeVideo(spec) → 미리 디코딩해 둔 <video>, isFollower(video), onAwaitPlay(video, spec) }
 */
(function(){
  function create(container, opts){
    opts = opts || {}
    let cur = null          // { key, video, spec, sig, started }
    let visible = false
    let hideTimer = null

    container.style.opacity = '0'
    container.style.zIndex = '0'

    const sigOf = spec => [spec.loop !== false, spec.loopMode || 'normal', spec.loopFade ?? 1].join('|')

    function startPlay(v, spec){
      // Chromium 자동재생 정책: 음소거로 먼저 재생한 뒤 소리 설정을 반영
      v.muted = true
      v.play().then(() => { v.muted = !spec.audio }).catch(() => {})
    }

    function attachLoop(v, spec){
      if (window.VideoLoop) {
        VideoLoop.attach(v, { loop: spec.loop !== false, mode: spec.loopMode, fade: spec.loopFade,
          isFollower: opts.isFollower ? () => opts.isFollower(v) : undefined })
      } else {
        v.loop = spec.loop !== false
      }
    }

    function makeVideo(spec){
      const v = (opts.takeVideo && opts.takeVideo(spec)) || document.createElement('video')
      if (!v.getAttribute('src')) v.src = spec.url
      v.playsInline = true
      v.preload = 'auto'
      v.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:' + (spec.fit || 'cover') +
        ';display:block;max-width:none;max-height:none;'
      return v
    }

    function removeVideo(v){
      if (!v) return
      if (window.VideoLoop) VideoLoop.detach(v)
      try { v.pause(); v.removeAttribute('src'); v.load() } catch (e) {}
      v.remove()
    }

    function clearHide(){
      if (hideTimer) { clearTimeout(hideTimer); hideTimer = null }
    }

    function show(spec, fd){
      if (!spec || !spec.url) { hide(fd); return null }
      fd = Math.max(0, parseFloat(fd) || 0)
      clearHide()
      const wasVisible = visible
      visible = true
      container.style.transition = 'none'
      container.style.zIndex = '0'
      container.style.opacity = '1'
      container.dataset.mId = spec.key || ''

      // 같은 영상: 그대로 이어서 재생
      if (cur && cur.key === spec.key) {
        if (cur.sig !== sigOf(spec)) { attachLoop(cur.video, spec); cur.sig = sigOf(spec) }
        cur.spec = spec
        cur.video.style.objectFit = spec.fit || 'cover'
        if (cur.video.paused && !cur.video.ended && (spec.autoPlay !== false || cur.started)) startPlay(cur.video, spec)
        else cur.video.muted = !spec.audio
        return cur.video
      }

      // 새 영상
      const old = cur
      const v = makeVideo(spec)
      cur = { key: spec.key, video: v, spec, sig: sigOf(spec), started: false }
      container.appendChild(v)
      attachLoop(v, spec)
      if (spec.autoPlay !== false) {
        cur.started = true
        startPlay(v, spec)
      } else {
        v.muted = !spec.audio
        v.addEventListener('play', () => { if (cur && cur.video === v) cur.started = true }, { once: true })
        if (opts.onAwaitPlay) opts.onAwaitPlay(v, spec)
      }

      if (old) {
        old.video.muted = true
        if (wasVisible && fd > 0) {
          // 그룹이 바뀌어 다른 영상이 되면 새 영상이 위에서 서서히 나타난다
          v.style.opacity = '0'
          void v.offsetWidth
          v.style.transition = 'opacity ' + fd + 's ease-in-out'
          v.style.opacity = '1'
          setTimeout(() => removeVideo(old.video), fd * 1000 + 50)
        } else {
          removeVideo(old.video)
        }
      }
      return v
    }

    function hide(fd){
      if (!visible) return
      fd = Math.max(0, parseFloat(fd) || 0)
      visible = false
      clearHide()
      const v = cur && cur.video
      const done = () => {
        hideTimer = null
        container.style.transition = 'none'
        container.style.opacity = '0'
        container.style.zIndex = '0'
        if (v) { try { v.pause() } catch (e) {} }
      }
      if (fd > 0) {
        container.style.zIndex = '1'
        container.style.transition = 'none'
        container.style.opacity = '1'
        void container.offsetWidth
        container.style.transition = 'opacity ' + fd + 's ease-in-out'
        container.style.opacity = '0'
        hideTimer = setTimeout(done, fd * 1000 + 30)
      } else {
        done()
      }
    }

    return {
      show,
      hide,
      active: () => visible && !!cur,
      video: () => (visible && cur) ? cur.video : null,
      key: () => cur ? cur.key : null
    }
  }

  window.GroupBgLayer = { create }
})()
