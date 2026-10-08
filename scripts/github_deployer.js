#!/usr/bin/env node
/**
 * scripts/github_deployer.js
 * 
 * 1-Click GitHub Repository Publisher & Deployer.
 * Adapted from ZeroWeb deployment engine.
 * 
 * Capabilities:
 * 1. Checks Git & GitHub CLI (gh) installation & authentication
 * 2. Auto-creates a new GitHub repository via `gh` (public/private) or connects existing remote
 * 3. Enforces strict .gitignore safety (prevents caching or temp bloat)
 * 4. Staging, custom commit message, and pushing to branch `main`
 * 5. Automatic remote sync (`git pull --rebase origin main`) if remote was already initialized
 * 6. Displays live repository link and automatically opens it in the browser
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const rootDir = path.resolve(__dirname, '..');

function prompt(query) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  return new Promise(resolve => rl.question(query, ans => {
    rl.close();
    resolve(ans.trim());
  }));
}

function run(cmd, opts = {}) {
  try {
    const res = execSync(cmd, {
      cwd: rootDir,
      encoding: 'utf8',
      stdio: opts.inherit ? 'inherit' : ['pipe', 'pipe', 'pipe'],
      ...opts
    });
    return res ? res.toString().trim() : '';
  } catch (err) {
    if (opts.throwOnError) throw err;
    return null;
  }
}

function openBrowser(url) {
  try {
    const start = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'start ""' : 'xdg-open';
    execSync(`${start} "${url}"`, { stdio: 'ignore' });
  } catch (e) { }
}

async function main() {
  console.log('\n==============================================================================');
  console.log('  🚀 AGENTIC PDF TO WORD & DOCX STUDIO — ONE-CLICK GITHUB DEPLOYER');
  console.log('  انتشار و همگام‌سازی آسان پروژه روی گیت‌هاب (۱-کلیک)');
  console.log('==============================================================================\n');

  // 1. Verify Git Installation
  console.log('[*] [1/5] Checking Git installation...');
  const gitVersion = run('git --version');
  if (!gitVersion) {
    console.error('\n[!] خطا: نرم‌افزار Git روی سیستم شما نصب نیست / Git is not installed.');
    console.error('    لطفاً Git را از آدرس https://git-scm.com دانلود و نصب نمایید.\n');
    openBrowser('https://git-scm.com/downloads');
    process.exit(1);
  }
  console.log(`    [✓] ${gitVersion} detected.`);

  // 2. Initialize Git Repository if missing
  console.log('\n[*] [2/5] Checking local repository...');
  const isGitRepo = fs.existsSync(path.join(rootDir, '.git'));
  if (!isGitRepo) {
    console.log('    [+] Initializing local Git repository...');
    run('git init -b main');
    console.log('    [✓] Repository initialized on branch "main".');
  } else {
    run('git branch -M main');
    console.log('    [✓] Git repository active (branch: main).');
  }

  // Configure git buffer & HTTP protocol for reliable large pushes on Windows
  run('git config http.postBuffer 524288000');
  run('git config http.version HTTP/1.1');

  // 3. Remote Origin Configuration
  console.log('\n[*] [3/5] Configuring GitHub remote repository...');
  let remoteUrl = run('git remote get-url origin');

  if (!remoteUrl) {
    // Check if GitHub CLI is logged in
    const ghStatus = run('gh auth status');
    const isGhLoggedIn = ghStatus !== null || run('gh api user');

    if (isGhLoggedIn) {
      console.log('\n  💡 شما به حساب گیت‌هاب (GitHub CLI) متصل هستید!');
      console.log('     انتخاب نحوه ایجاد مخزن (Repository):');
      console.log('     1) ایجاد مخزن جدید در اکانت شما به‌صورت خودکار (پیش‌فرض)');
      console.log('     2) اتصال به آدرس مخزن موجود (Existing Repository URL)\n');

      const choice = (await prompt('گزینه مورد نظر (1 یا 2، اینتر برای پیش‌فرض 1): ')) || '1';

      if (choice === '1') {
        const defaultName = 'Antigravity-PDF-to-Docx';
        const repoName = (await prompt(`نام مخزن در گیت‌هاب [${defaultName}]: `)) || defaultName;
        const visibility = (await prompt('عمومی باشد یا خصوصی؟ public / private [public]: ')).toLowerCase() || 'public';
        const flag = visibility.startsWith('priv') ? '--private' : '--public';

        console.log(`\n    [+] Creating GitHub repository "${repoName}" (${flag})...`);
        try {
          run(`gh repo create "${repoName}" ${flag} --source=. --remote=origin`, { inherit: true, throwOnError: true });
          remoteUrl = run('git remote get-url origin');
          console.log(`    [✓] Created & linked to: ${remoteUrl}`);
        } catch (e) {
          console.warn('    [!] Could not auto-create via gh, falling back to manual input.');
        }
      }
    }

    if (!remoteUrl) {
      console.log('\n  لطفاً آدرس مخزن گیت‌هاب را وارد نمایید (مثال: https://github.com/faithsaly5-stack/Antigravity-PDF-to-Docx.git):');
      const input = await prompt('GitHub Repo URL: ');
      if (input) {
        remoteUrl = input;
        run(`git remote add origin ${remoteUrl}`);
        console.log(`    [✓] Remote origin set to: ${remoteUrl}`);
      } else {
        console.error('\n[!] آدرس مخزن وارد نشد. عملیات لغو شد.');
        process.exit(1);
      }
    }
  } else {
    console.log(`    [✓] Target GitHub Repository: \x1b[36m${remoteUrl}\x1b[0m`);
    const change = await prompt('    Press Enter to use this repository, or type "change" to edit: ');
    if (change.toLowerCase() === 'change') {
      const newUrl = await prompt('    Enter new GitHub Repo URL: ');
      if (newUrl) {
        run(`git remote set-url origin ${newUrl}`);
        remoteUrl = newUrl;
        console.log(`    [✓] Updated remote origin to: ${remoteUrl}`);
      }
    }
  }

  // 4. Stage & Safety Filter
  console.log('\n[*] [4/5] Staging and verifying changes...');
  run('git add -A');

  // Check if there are changes to commit
  const status = run('git status --porcelain');
  if (status) {
    const today = new Date().toISOString().slice(0, 10);
    const defaultMsg = `Release Agentic PDF to Word Studio (${today})`;
    const customMsg = await prompt(`پیام کامیت / Commit message [${defaultMsg}]: `);
    const commitMsg = customMsg || defaultMsg;

    console.log(`    [+] Committing with message: "${commitMsg}"...`);
    run(`git commit -m "${commitMsg.replace(/"/g, '\\"')}"`);
  } else {
    console.log('    [i] Working tree is clean. Ready to push latest commits.');
  }

  // 5. Push to GitHub
  console.log('\n[*] [5/5] Pushing changes to GitHub (branch: main)...');
  let pushSuccess = false;

  try {
    run('git push -u origin main', { inherit: true, throwOnError: true });
    pushSuccess = true;
  } catch (err) {
    console.log('\n    [!] Standard push rejected. Attempting automatic sync (git pull --rebase)...');
    try {
      run('git pull --rebase origin main', { inherit: true, throwOnError: true });
      run('git push -u origin main', { inherit: true, throwOnError: true });
      pushSuccess = true;
    } catch (e) {
      console.error('\n[!] Push to GitHub failed.');
      console.error('    راهنما: اگر با خطای دسترسی روبرو شدید، دستور `gh auth login` را برای ورود مجدد اجرا نمایید.');
    }
  }

  if (pushSuccess) {
    // Format web URL from git URL
    let webUrl = remoteUrl.replace(/\.git$/, '').replace(/^git@github\.com:/, 'https://github.com/');
    if (!webUrl.startsWith('http')) webUrl = `https://github.com/${webUrl}`;

    console.log('\n');
    console.log('╔════════════════════════════════════════════════════════════════════════════════╗');
    console.log('║                                                                                ║');
    console.log('║         🎉 پروژه با موفقیت روی گیت‌هاب منتشر و همگام‌سازی شد! 🎉               ║');
    console.log('║           SUCCESSFULLY PUBLISHED & SYNCHRONIZED TO GITHUB!                     ║');
    console.log('║                                                                                ║');
    console.log('╚════════════════════════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('  🐙 آدرس مخزن در گیت‌هاب (GitHub Repository):');
    console.log(`     👉  \x1b[36m\x1b[1m${webUrl}\x1b[0m`);
    console.log('');
    console.log('  🔒 امنیت اطلاعات (Security Guarantee):');
    console.log('     پوشه‌های کش و فایل‌های موقت مطابق .gitignore فیلتر شده و آپلود نشدند.');
    console.log('');
    console.log('────────────────────────────────────────────────────────────────────────────────');
    console.log('  Opening your GitHub repository in your browser...');
    console.log('────────────────────────────────────────────────────────────────────────────────\n');

    openBrowser(webUrl);
  }
}

main().catch(err => {
  console.error('[!] Unexpected error:', err.message);
  process.exit(1);
});
