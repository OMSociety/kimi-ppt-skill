/**
 * Cordis host plugin for dsh-kimi-ppt-skill.
 *
 * One job: register the bundled `skills/` directory so DSH discovers the
 * `skills/kimi-ppt` skill. DSH finds skills one level deep under a fixed set of
 * roots, and a path inside `node_modules` is none of them — so shipping
 * `skills/` in the package file list puts the files on disk where nothing will
 * ever look at them. Registering the directory is what turns them into a skill.
 *
 * Best effort on purpose: the skill is the product, and a host that exposes no
 * skill-root hook should still load rather than failing.
 */
import { existsSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'dsh-kimi-ppt-skill'
export const inject = []

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const SKILLS_DIR = path.join(ROOT, 'skills')

function registerSkillRoot(ctx) {
  if (!existsSync(SKILLS_DIR)) return () => {}
  for (const register of [ctx.skills?.addRoot, ctx.skills?.register, ctx.addSkillRoot]) {
    if (typeof register !== 'function') continue
    try {
      const off = register.call(ctx.skills ?? ctx, SKILLS_DIR)
      return typeof off === 'function' ? off : () => {}
    } catch (err) {
      ctx.logger?.warn?.(`[dsh-kimi-ppt-skill] could not register the skill directory: ${err.message}`)
      return () => {}
    }
  }
  ctx.logger?.info?.(
    `[dsh-kimi-ppt-skill] this host exposes no skill-root hook; to use the skill, link ${SKILLS_DIR} into a DSH skills directory`
  )
  return () => {}
}

export function apply(ctx) {
  const unregisterSkills = registerSkillRoot(ctx)
  ctx.effect(() => () => {
    unregisterSkills()
  })
}
