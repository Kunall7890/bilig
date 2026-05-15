import * as THREE from 'https://esm.sh/three@0.184.0'

const canvas = document.querySelector('#hero-workbook-scene')
const media = canvas?.closest('.hero-media')

const sheetSize = {
  height: 3,
  width: 5.2,
}

const grid = {
  height: 1.9,
  headerHeight: 0.32,
  rowHeaderWidth: 0.38,
  top: 0.62,
  width: 4.52,
  x: -2.26,
}

const columns = 4
const rows = 5
const cellWidth = (grid.width - grid.rowHeaderWidth) / columns
const rowHeight = (grid.height - grid.headerHeight) / rows
const cellTop = grid.top - grid.headerHeight

if (canvas instanceof HTMLCanvasElement && media instanceof HTMLElement) {
  const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches
  const renderer = new THREE.WebGLRenderer({
    alpha: true,
    antialias: true,
    canvas,
    powerPreference: 'high-performance',
  })
  renderer.outputColorSpace = THREE.SRGBColorSpace
  renderer.setClearColor(0x000000, 0)
  renderer.shadowMap.enabled = true
  renderer.shadowMap.type = THREE.PCFShadowMap
  renderer.toneMapping = THREE.ACESFilmicToneMapping
  renderer.toneMappingExposure = 1.05

  const scene = new THREE.Scene()
  const camera = new THREE.PerspectiveCamera(34, 1, 0.1, 100)
  camera.position.set(0.16, 0.38, 7.8)
  camera.lookAt(0.02, -0.06, 0)

  const root = new THREE.Group()
  root.position.set(-0.18, -0.04, 0)
  root.rotation.set(-0.5, 0.12, -0.15)
  scene.add(root)

  scene.add(new THREE.HemisphereLight(0xfbfff7, 0x020503, 1.08))

  const keyLight = new THREE.DirectionalLight(0xffffff, 3.6)
  keyLight.position.set(-2.9, 4.1, 5.6)
  keyLight.castShadow = true
  keyLight.shadow.mapSize.set(2048, 2048)
  keyLight.shadow.camera.near = 0.5
  keyLight.shadow.camera.far = 18
  keyLight.shadow.camera.left = -5.5
  keyLight.shadow.camera.right = 5.5
  keyLight.shadow.camera.top = 5.5
  keyLight.shadow.camera.bottom = -5.5
  scene.add(keyLight)

  const rimLight = new THREE.DirectionalLight(0x64ff9b, 2.2)
  rimLight.position.set(4.2, 1.8, 3.3)
  scene.add(rimLight)

  const fillLight = new THREE.PointLight(0x29d474, 3.8, 7)
  fillLight.position.set(0.4, -0.65, 2.2)
  scene.add(fillLight)

  const shadowPlane = new THREE.Mesh(
    new THREE.PlaneGeometry(7.2, 4.8),
    new THREE.MeshBasicMaterial({
      map: makeRadialTexture(1024),
      opacity: 0.78,
      transparent: true,
      depthWrite: false,
    }),
  )
  shadowPlane.position.set(0.1, -0.12, -0.5)
  root.add(shadowPlane)

  buildWorkbookArtifact(root)
  const route = addFormulaRoute(root)
  const particles = makeParticles()
  scene.add(particles)

  const startedAt = performance.now()
  let animationFrame = 0

  function resize() {
    const rect = canvas.getBoundingClientRect()
    if (rect.width <= 0 || rect.height <= 0) {
      return
    }
    const dpr = Math.min(window.devicePixelRatio || 1, 2)
    renderer.setPixelRatio(dpr)
    renderer.setSize(rect.width, rect.height, false)
    camera.aspect = rect.width / rect.height
    camera.position.z = camera.aspect < 1 ? 8.9 : 7.6
    root.scale.setScalar(camera.aspect < 1 ? 0.62 : 0.66)
    camera.updateProjectionMatrix()
  }

  function render() {
    resize()
    renderer.render(scene, camera)
  }

  function animate() {
    const elapsed = (performance.now() - startedAt) / 1000
    root.rotation.x = -0.5 + Math.sin(elapsed * 0.22) * 0.014
    root.rotation.y = 0.12 + Math.sin(elapsed * 0.27) * 0.014
    root.rotation.z = -0.15 + Math.sin(elapsed * 0.19) * 0.008
    root.position.y = -0.04 + Math.sin(elapsed * 0.36) * 0.02

    const pulse = 1 + Math.sin(elapsed * 1.55) * 0.045
    route.inputNode.scale.setScalar(pulse)
    route.resultNode.scale.setScalar(1 + Math.sin(elapsed * 1.55 + 0.7) * 0.06)
    route.path.material.emissiveIntensity = 0.74 + Math.sin(elapsed * 1.2) * 0.12
    route.signal.position.copy(route.curve.getPointAt((elapsed * 0.16) % 1))
    route.signal.scale.setScalar(0.82 + Math.sin(elapsed * 3.1) * 0.12)
    fillLight.intensity = 3.55 + Math.sin(elapsed * 1.05) * 0.34
    particles.rotation.z = elapsed * 0.008

    renderer.render(scene, camera)
    animationFrame = window.requestAnimationFrame(animate)
  }

  const observer = new ResizeObserver(render)
  observer.observe(media)
  window.addEventListener('orientationchange', render)

  media.classList.add('is-scene-ready')
  render()
  if (!reducedMotion) {
    animationFrame = window.requestAnimationFrame(animate)
  }

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      window.cancelAnimationFrame(animationFrame)
      return
    }
    if (!reducedMotion) {
      animationFrame = window.requestAnimationFrame(animate)
    }
  })
}

function buildWorkbookArtifact(group) {
  const baseMaterial = new THREE.MeshPhysicalMaterial({
    color: 0x030806,
    clearcoat: 0.72,
    clearcoatRoughness: 0.18,
    metalness: 0.18,
    roughness: 0.28,
  })
  const shadowMaterial = new THREE.MeshBasicMaterial({
    color: 0x092415,
    opacity: 0.42,
    transparent: true,
  })
  const paperMaterial = new THREE.MeshPhysicalMaterial({
    color: 0xfbfcf7,
    clearcoat: 0.38,
    clearcoatRoughness: 0.28,
    metalness: 0.01,
    roughness: 0.46,
  })
  const glassMaterial = new THREE.MeshPhysicalMaterial({
    color: 0xe9f3e7,
    opacity: 0.54,
    roughness: 0.18,
    transparent: true,
    transmission: 0.08,
  })
  const headerMaterial = new THREE.MeshBasicMaterial({
    color: 0xe8f1e5,
    opacity: 0.44,
    transparent: true,
  })
  const activeMaterial = new THREE.MeshPhysicalMaterial({
    color: 0xaaf2c5,
    emissive: 0x145f3d,
    emissiveIntensity: 0.14,
    opacity: 0.82,
    roughness: 0.22,
    transparent: true,
  })
  const accentMaterial = new THREE.MeshBasicMaterial({
    color: 0x168552,
    opacity: 0.22,
    transparent: true,
  })

  const backGlow = makeRoundedBox(5.85, 3.62, 0.08, 0.22, shadowMaterial)
  backGlow.position.set(0.14, -0.16, -0.34)
  group.add(backGlow)

  const base = makeRoundedBox(5.64, 3.44, 0.24, 0.18, baseMaterial)
  base.position.z = -0.15
  base.castShadow = true
  base.receiveShadow = true
  group.add(base)

  const lowerSheet = makeRoundedBox(sheetSize.width * 0.98, sheetSize.height * 0.96, 0.035, 0.12, glassMaterial)
  lowerSheet.position.set(-0.14, -0.15, -0.035)
  lowerSheet.rotation.z = 0.018
  lowerSheet.receiveShadow = true
  group.add(lowerSheet)

  const paper = makeRoundedBox(sheetSize.width, sheetSize.height, 0.055, 0.12, paperMaterial)
  paper.position.set(0.04, 0.03, 0.045)
  paper.castShadow = true
  paper.receiveShadow = true
  group.add(paper)

  group.add(
    makeSurfacePanel(grid.width, grid.headerHeight, grid.x + grid.width / 2, grid.top - grid.headerHeight / 2, 0.122, headerMaterial),
    makeSurfacePanel(grid.rowHeaderWidth, grid.height, grid.x + grid.rowHeaderWidth / 2, grid.top - grid.height / 2, 0.123, headerMaterial),
  )

  const formulaBar = makeRoundedBox(2.72, 0.26, 0.07, 0.08, glassMaterial)
  formulaBar.position.set(-0.64, 1.16, 0.28)
  formulaBar.castShadow = true
  formulaBar.receiveShadow = true
  group.add(formulaBar)

  const formulaText = makeFormulaLabel()
  formulaText.position.set(-0.64, 1.16, 0.322)
  group.add(formulaText)

  group.add(makeGridLines())
  addCellAccents(group, accentMaterial)

  const input = makeCellTile(1, 1, activeMaterial)
  const output = makeCellTile(3, 4, activeMaterial)
  input.position.z = 0.17
  output.position.z = 0.172
  input.castShadow = true
  output.castShadow = true
  group.add(input, output)

  const overlay = new THREE.Mesh(
    new THREE.PlaneGeometry(sheetSize.width, sheetSize.height),
    new THREE.MeshBasicMaterial({
      map: makeSheetLabelTexture(),
      transparent: true,
      depthWrite: false,
    }),
  )
  overlay.position.z = 0.19
  group.add(overlay)

  const sheen = new THREE.Mesh(
    new THREE.PlaneGeometry(sheetSize.width * 0.95, sheetSize.height * 0.82),
    new THREE.MeshBasicMaterial({
      color: 0xffffff,
      opacity: 0.08,
      transparent: true,
      depthWrite: false,
    }),
  )
  sheen.position.set(-0.18, 0.08, 0.205)
  sheen.rotation.z = -0.04
  group.add(sheen)
}

function addFormulaRoute(group) {
  const material = new THREE.MeshStandardMaterial({
    color: 0x46df85,
    emissive: 0x1fab60,
    emissiveIntensity: 0.82,
    metalness: 0.06,
    roughness: 0.14,
  })
  const nodeMaterial = new THREE.MeshPhysicalMaterial({
    color: 0x47e989,
    clearcoat: 0.7,
    emissive: 0x139e55,
    emissiveIntensity: 0.64,
    metalness: 0.04,
    roughness: 0.18,
  })
  const glowMaterial = new THREE.MeshBasicMaterial({
    blending: THREE.AdditiveBlending,
    color: 0x49f28a,
    depthWrite: false,
    opacity: 0.16,
    transparent: true,
  })

  const start = cellPosition(1, 1, 0.28)
  const end = cellPosition(3, 4, 0.3)
  const curve = new THREE.CatmullRomCurve3([start, new THREE.Vector3(-0.44, 0.38, 0.62), new THREE.Vector3(0.86, -0.1, 0.64), end])

  const path = new THREE.Mesh(new THREE.TubeGeometry(curve, 112, 0.024, 20, false), material)
  path.castShadow = true
  group.add(path)

  const inputNode = makeNode(nodeMaterial, start, 0.085)
  const resultNode = makeNode(nodeMaterial, end, 0.11)
  const signal = makeNode(nodeMaterial, start, 0.052)
  signal.castShadow = true
  group.add(inputNode, resultNode, signal)
  group.add(makeGlow(glowMaterial, start, 0.22), makeGlow(glowMaterial, end, 0.3))

  return {
    curve,
    inputNode,
    path,
    resultNode,
    signal,
  }
}

function makeCellTile(column, row, material) {
  const tile = makeRoundedBox(cellWidth * 0.82, rowHeight * 0.72, 0.04, 0.045, material)
  const position = cellPosition(column, row, 0)
  tile.position.set(position.x, position.y, 0)
  return tile
}

function cellPosition(column, row, z) {
  return new THREE.Vector3(grid.x + grid.rowHeaderWidth + cellWidth * (column + 0.5), cellTop - rowHeight * (row + 0.5), z)
}

function makeGridLines() {
  const positions = []
  const left = grid.x
  const right = grid.x + grid.width
  const bottom = grid.top - grid.height

  for (let index = 0; index <= columns; index += 1) {
    const x = grid.x + grid.rowHeaderWidth + cellWidth * index
    positions.push(x, bottom, 0.132, x, grid.top, 0.132)
  }
  for (let index = 0; index <= rows; index += 1) {
    const y = cellTop - rowHeight * index
    positions.push(left, y, 0.132, right, y, 0.132)
  }
  positions.push(left, grid.top, 0.132, right, grid.top, 0.132)
  positions.push(left, bottom, 0.132, right, bottom, 0.132)
  positions.push(left, bottom, 0.132, left, grid.top, 0.132)
  positions.push(right, bottom, 0.132, right, grid.top, 0.132)

  const geometry = new THREE.BufferGeometry()
  geometry.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3))
  return new THREE.LineSegments(
    geometry,
    new THREE.LineBasicMaterial({
      color: 0x8dab86,
      opacity: 0.2,
      transparent: true,
    }),
  )
}

function addCellAccents(group, material) {
  const bars = [
    [0, 1, 0.42],
    [2, 1, 0.48],
    [3, 1, 0.54],
    [1, 2, 0.36],
    [2, 2, 0.34],
    [3, 2, 0.4],
    [1, 3, 0.3],
    [2, 3, 0.36],
    [3, 3, 0.32],
  ]

  for (const [column, row, width] of bars) {
    const position = cellPosition(column, row, 0.205)
    const bar = makeRoundedBox(width, 0.026, 0.012, 0.012, material)
    bar.position.set(position.x + cellWidth * 0.18, position.y - rowHeight * 0.12, position.z)
    group.add(bar)
  }
}

function makeSurfacePanel(width, height, x, y, z, material) {
  const panel = new THREE.Mesh(new THREE.PlaneGeometry(width, height), material)
  panel.position.set(x, y, z)
  return panel
}

function makeRoundedBox(width, height, depth, radius, material) {
  const shape = makeRoundedRectShape(width, height, radius)
  const geometry = new THREE.ExtrudeGeometry(shape, {
    bevelEnabled: true,
    bevelSegments: 8,
    bevelSize: Math.min(depth * 0.38, radius * 0.42),
    bevelThickness: Math.min(depth * 0.35, radius * 0.32),
    curveSegments: 12,
    depth,
  })
  geometry.center()
  const mesh = new THREE.Mesh(geometry, material)
  mesh.castShadow = true
  mesh.receiveShadow = true
  return mesh
}

function makeRoundedRectShape(width, height, radius) {
  const x = -width / 2
  const y = -height / 2
  const shape = new THREE.Shape()
  shape.moveTo(x + radius, y)
  shape.lineTo(x + width - radius, y)
  shape.quadraticCurveTo(x + width, y, x + width, y + radius)
  shape.lineTo(x + width, y + height - radius)
  shape.quadraticCurveTo(x + width, y + height, x + width - radius, y + height)
  shape.lineTo(x + radius, y + height)
  shape.quadraticCurveTo(x, y + height, x, y + height - radius)
  shape.lineTo(x, y + radius)
  shape.quadraticCurveTo(x, y, x + radius, y)
  return shape
}

function makeNode(material, position, radius) {
  const node = new THREE.Mesh(new THREE.SphereGeometry(radius, 32, 32), material)
  node.position.copy(position)
  node.castShadow = true
  node.receiveShadow = true
  return node
}

function makeGlow(material, position, radius) {
  const glow = new THREE.Mesh(new THREE.SphereGeometry(radius, 36, 36), material)
  glow.position.copy(position)
  return glow
}

function makeFormulaLabel() {
  const textureCanvas = document.createElement('canvas')
  textureCanvas.width = 900
  textureCanvas.height = 120
  const context = getCanvasContext(textureCanvas)
  context.clearRect(0, 0, textureCanvas.width, textureCanvas.height)
  drawPill(context, 28, 34, 92, 52, 14, '#e3ece1', '#c6d6c2')
  drawText(context, 'D5', 74, 62, 24, '#5a6659', 760, 'center', true)
  drawText(context, '=SUM(D2:D4)', 158, 62, 27, '#127a47', 760, 'left', true)
  const texture = new THREE.CanvasTexture(textureCanvas)
  texture.colorSpace = THREE.SRGBColorSpace
  texture.needsUpdate = true

  return new THREE.Mesh(
    new THREE.PlaneGeometry(2.54, 0.24),
    new THREE.MeshBasicMaterial({
      map: texture,
      transparent: true,
      depthWrite: false,
    }),
  )
}

function makeSheetLabelTexture() {
  const textureCanvas = document.createElement('canvas')
  textureCanvas.width = 1800
  textureCanvas.height = 1040
  const context = getCanvasContext(textureCanvas)
  context.clearRect(0, 0, textureCanvas.width, textureCanvas.height)

  for (const [index, column] of ['A', 'B', 'C', 'D'].entries()) {
    const x = grid.x + grid.rowHeaderWidth + cellWidth * (index + 0.5)
    drawSceneText(context, column, x, grid.top - 0.16, 26, 'rgba(92, 115, 87, 0.62)', 740)
  }
  for (let row = 1; row <= rows; row += 1) {
    drawSceneText(context, String(row), grid.x + 0.19, cellTop - rowHeight * (row - 0.5), 22, 'rgba(112, 137, 105, 0.5)', 720)
  }

  drawSceneText(context, 'Customers', cellPosition(1, 0, 0).x, cellPosition(1, 0, 0).y, 22, 'rgba(24, 36, 23, 0.72)', 700)
  drawSceneText(context, 'Revenue', cellPosition(3, 0, 0).x, cellPosition(3, 0, 0).y, 22, 'rgba(24, 36, 23, 0.72)', 700)
  drawSceneText(context, '32', cellPosition(1, 1, 0).x, cellPosition(1, 1, 0).y, 35, '#12804d', 780, 'center', true)
  drawSceneText(context, '51,300', cellPosition(3, 4, 0).x, cellPosition(3, 4, 0).y, 35, '#12804d', 780, 'center', true)

  const texture = new THREE.CanvasTexture(textureCanvas)
  texture.colorSpace = THREE.SRGBColorSpace
  texture.needsUpdate = true
  return texture

  function drawSceneText(drawContext, text, sceneX, sceneY, size, color, weight, align = 'center', mono = false) {
    drawText(drawContext, text, mapX(sceneX), mapY(sceneY), size, color, weight, align, mono)
  }

  function mapX(x) {
    return ((x + sheetSize.width / 2) / sheetSize.width) * textureCanvas.width
  }

  function mapY(y) {
    return ((sheetSize.height / 2 - y) / sheetSize.height) * textureCanvas.height
  }
}

function makeParticles() {
  const particles = new THREE.Group()
  let seed = 97
  for (let index = 0; index < 22; index += 1) {
    const dot = new THREE.Mesh(
      new THREE.SphereGeometry(0.01 + seededRandom() * 0.012, 8, 8),
      new THREE.MeshBasicMaterial({
        color: index % 6 === 0 ? 0x35d179 : 0xdfe8dd,
        opacity: index % 6 === 0 ? 0.18 : 0.1,
        transparent: true,
      }),
    )
    dot.position.set((seededRandom() - 0.5) * 7.3, (seededRandom() - 0.5) * 4.4, -1 - seededRandom() * 1.5)
    particles.add(dot)
  }
  return particles

  function seededRandom() {
    seed = (seed * 1664525 + 1013904223) % 4294967296
    return seed / 4294967296
  }
}

function makeRadialTexture(size) {
  const textureCanvas = document.createElement('canvas')
  textureCanvas.width = size
  textureCanvas.height = size
  const context = getCanvasContext(textureCanvas)
  const gradient = context.createRadialGradient(size / 2, size / 2, 0, size / 2, size / 2, size / 2)
  gradient.addColorStop(0, 'rgba(53, 209, 121, 0.42)')
  gradient.addColorStop(0.42, 'rgba(18, 58, 35, 0.22)')
  gradient.addColorStop(1, 'rgba(0, 0, 0, 0)')
  context.fillStyle = gradient
  context.fillRect(0, 0, size, size)
  const texture = new THREE.CanvasTexture(textureCanvas)
  texture.colorSpace = THREE.SRGBColorSpace
  texture.needsUpdate = true
  return texture
}

function getCanvasContext(textureCanvas) {
  const context = textureCanvas.getContext('2d')
  if (context === null) {
    throw new Error('Canvas 2D context is unavailable')
  }
  return context
}

function drawPill(context, x, y, width, height, radius, fill, stroke) {
  context.beginPath()
  context.moveTo(x + radius, y)
  context.lineTo(x + width - radius, y)
  context.quadraticCurveTo(x + width, y, x + width, y + radius)
  context.lineTo(x + width, y + height - radius)
  context.quadraticCurveTo(x + width, y + height, x + width - radius, y + height)
  context.lineTo(x + radius, y + height)
  context.quadraticCurveTo(x, y + height, x, y + height - radius)
  context.lineTo(x, y + radius)
  context.quadraticCurveTo(x, y, x + radius, y)
  context.closePath()
  context.fillStyle = fill
  context.strokeStyle = stroke
  context.lineWidth = 2
  context.fill()
  context.stroke()
}

function drawText(context, text, x, y, size, color, weight, align, mono = false) {
  context.fillStyle = color
  context.font = `${weight} ${size}px ${
    mono ? 'SFMono-Regular, Menlo, ui-monospace, monospace' : 'Inter, ui-sans-serif, system-ui, sans-serif'
  }`
  context.textAlign = align
  context.textBaseline = 'middle'
  context.fillText(text, x, y)
}
