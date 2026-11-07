/**
 * 그룹 배경 영상 레이어 (편집 미리보기 index.html 과 송출 화면 screen.html 이 함께 쓴다)
 *
 * 그룹 배경이 영상이면, 그 영상을 그대로 쓰는 슬라이드마다 <video>를 새로 만들지 않고
 * A/B 슬라이드 레이어 뒤에 있는 이 레이어의 영상 하나를 계속 이어서 재생한다.
 * 슬라이드를 넘겨도 영상은 처음부터 다시 틀지 않고, 위의 가사·텍스트만 바뀐다.
 *
 *  show(spec, fd) : 이 영상을 보여준다. 이미 같은 영상이면 그대로 이어서 재생한다.
 *                   다른 영상이면 새 영상의 첫 화면이 준비된 뒤 fd초 동안 서서히 나타난다.
 *  hide(fd)       : 따로 배경을 넣은 슬라이드로 넘어갈 때. fd초 동안 사라진 뒤 일시정지한다.
 *                   (다시 show 하면 멈춘 위치부터 이어서 재생)
 *
 * z-index: 평소 0 (두 슬라이드 레이어 아래). 사라지는 동안만 1로 올려, 같은 z-index 1인
 *          '다음 슬라이드 레이어' 위에서 함께 페이드아웃되게 한다. (컨테이너는 LB 뒤에 위치해야 함)
 *
 * spec = { key, url, fit, audio, loop, loopMode, loopFade, autoPlay }
 * opts = { takeVideo(spec) → 미리 디코딩해 둔 <video>, isFollower(video), onAwaitPlay(video, spec),
 *          seamless: false 면 보통 반복은 브라우저 기본 반복 (편집 미리보기: 디코더를 아낀다) }
 *
 * 반복 재생은 VideoLoop의 이어달리기를 쓰므로 영상 하나가 <video> 두 개로 번갈아 재생될 수 있다.
 * cur.video 는 그 영상을 대표하는 요소이고, 지금 재생 중인 요소는 act(cur.video)로 얻는다.
 */
(function(){
  /** 지금 보이며 재생 중인 요소 */
  const act = v => (window.VideoLoop && v) ? VideoLoop.active(v) : v
  /** 이 영상을 이루는 요소들 (이어달리기 쌍둥이 포함) */
  const all = v => (window.VideoLoop && v) ? VideoLoop.elements(v) : (v ? [v] : [])

  function create(container, opts){
    opts = opts || {}
    let cur = null          // { key, video, spec, sig, started }
    let visible = false
    let hideTimer = null
    let retiring = []       // 새 영상이 다 나타나면 지울 이전 영상들 (그동안 아래에서 계속 재생된다)

    container.style.opacity = '0'
    container.style.zIndex = '0'

    const sigOf = spec => [spec.loop !== false, spec.loopMode || 'normal', spec.loopFade ?? 1].join('|')

    function startPlay(v, spec){
      // Chromium 자동재생 정책: 음소거로 먼저 재생한 뒤 소리 설정을 반영
      v = act(v)
      v.muted = true
      v.play().then(() => { v.muted = !spec.audio }).catch(() => {})
    }

    /** 반복 방식을 건다. 앞으로 이 영상을 대표할 요소를 돌려준다 */
    function attachLoop(v, spec){
      if (window.VideoLoop) {
        return VideoLoop.attach(v, { loop: spec.loop !== false, mode: spec.loopMode, fade: spec.loopFade,
          seamless: opts.seamless, isFollower: opts.isFollower ? () => opts.isFollower(act(v)) : undefined })
      }
      v.loop = spec.loop !== false
      return v
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
      // 이어달리기 쌍둥이까지 함께 내린다
      const keep = window.VideoLoop ? VideoLoop.detach(v) : v
      ;[v, keep].forEach(x => { if (!x) return; try { x.pause(); x.removeAttribute('src'); x.load() } catch (e) {} ; x.remove() })
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
        if (cur.sig !== sigOf(spec)) { cur.video = attachLoop(cur.video, spec); cur.sig = sigOf(spec) }
        cur.spec = spec
        all(cur.video).forEach(x => { x.style.objectFit = spec.fit || 'cover' })
        const a = act(cur.video)
        if (a.paused && !a.ended && (spec.autoPlay !== false || cur.started)) startPlay(a, spec)
        else a.muted = !spec.audio
        return a
      }

      // 새 영상
      const old = cur
      let v = makeVideo(spec)
      container.appendChild(v)
      v = attachLoop(v, spec)
      cur = { key: spec.key, video: v, spec, sig: sigOf(spec), started: false }
      if (spec.autoPlay !== false) {
        cur.started = true
        startPlay(v, spec)
      } else {
        v.muted = !spec.audio
        v.addEventListener('play', () => { if (cur && cur.video === v) cur.started = true }, { once: true })
        if (opts.onAwaitPlay) opts.onAwaitPlay(v, spec)
      }

      if (old) {
        all(old.video).forEach(x => { x.muted = true })
        if (wasVisible && fd > 0) {
          // 다른 영상으로 바뀌면 새 영상이 위에서 서서히 나타난다.
          // 새 영상의 첫 화면이 준비된 뒤에 시작해야 중간에 툭 튀어나오지 않는다 (최대 0.7초 대기).
          // 그동안 이전 영상은 아래에서 계속 재생되고, 새 영상이 다 나타나면 정리한다.
          retiring.push(old.video)
          v.style.opacity = '0'
          const ready = (window.VideoLoop && VideoLoop.whenFrame) ? VideoLoop.whenFrame(v, 700) : Promise.resolve()
          ready.then(() => {
            if (!cur || cur.video !== v) return   // 그 사이 또 다른 영상으로 바뀜 → 그쪽이 이어받아 정리한다
            v.style.transition = 'opacity ' + fd + 's ease-in-out'
            void v.offsetWidth
            v.style.opacity = '1'
            const olds = retiring
            retiring = []
            setTimeout(() => olds.forEach(removeVideo), fd * 1000 + 50)
          })
        } else {
          removeVideo(old.video)
          retiring.forEach(removeVideo)
          retiring = []
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
        if (v) { try { act(v).pause() } catch (e) {} }
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
      /** 지금 보이며 재생 중인 요소 */
      video: () => (visible && cur) ? act(cur.video) : null,
      key: () => cur ? cur.key : null
    }
  }

  window.GroupBgLayer = { create }
})()
