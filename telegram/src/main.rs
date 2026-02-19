use anyhow::Result as AResult;
use chrono::{Datelike, Local};
use dotenvy::dotenv;
use engine::excel::get_filled_table;
use std::{env, error::Error, net::SocketAddr};
use teloxide::{
    dispatching::dialogue::InMemStorage, prelude::*, types::InputFile, update_listeners::webhooks,
};

type UserDialogue = Dialogue<DState, InMemStorage<DState>>;

//Dialogue State
#[derive(Clone, Default)]
enum DState {
    #[default]
    Start,
    Salary,
}

#[tokio::main]
async fn main() {
    // Load logs
    pretty_env_logger::init();
    log::info!("Запуск бота...");
    // Load envs
    log::info!("Загрузка env...");
    dotenv().ok();
    // ();
    let token = env::var("BOT_TOKEN").expect("Не найден токен бота в .env файле!");
    let port = env::var("PORT").unwrap_or_else(|_| {
        log::error!("Порт не указан! Используем 8080!");
        "8080".to_string()
    });
    let url = env::var("WEBHOOK_URL").expect("Не найден WEBHOOK_URL в .env файле!");
    run_bot(token, port, url).await;
}

async fn run_bot(token: String, port: String, webhook_url: String) {
    // Init bot
    let bot = Bot::new(token);
    // Init ipv4 addr
    let addr = SocketAddr::from(([0, 0, 0, 0], port.parse::<u16>().unwrap()));
    // Setup webhook listener
    let listener = webhooks::axum(
        bot.clone(),
        webhooks::Options::new(addr, webhook_url.parse().unwrap()),
    )
    .await
    .expect("Не удалось поднять Webhook!");
    // Dialogue update logic
    let router = Update::filter_message()
        .enter_dialogue::<Message, InMemStorage<DState>, DState>()
        .branch(dptree::case![DState::Start].endpoint(start))
        .branch(dptree::case![DState::Salary].endpoint(salary));
    // Dispatcher
    Dispatcher::builder(bot, router)
        .dependencies(dptree::deps![InMemStorage::<DState>::new()])
        .enable_ctrlc_handler()
        .build()
        .dispatch_with_listener(
            listener,
            LoggingErrorHandler::with_custom_text("Ошибка при обновлении!"),
        )
        .await;
}

async fn start(bot: Bot, dialogue: UserDialogue, msg: Message) -> AResult<()> {
    let user = msg.from;
    let user_name: String = match user {
        Some(user) => match user.username {
            Some(username) => username,
            None => user.id.0.to_string(),
        },
        None => "пользователь".to_string(),
    };
    bot.send_message(
        msg.chat.id,
        format!(
            "Привет {}!\nВведи свой оклад ниже что бы получить готовый табель за {} год.",
            user_name,
            Local::now().year(),
        ),
    )
    .await?;
    dialogue.update(DState::Salary).await?;
    Ok(())
}

async fn salary(bot: Bot, msg: Message) -> AResult<()> {
    match msg.text() {
        Some(text) => {
            let salary = text.parse::<u32>().ok();
            match salary {
                Some(s) => {
                    let table = get_filled_table(s).await?;
                    bot.send_document(msg.chat.id, InputFile::memory(table))
                        .await?;
                }
                None => todo!(),
            };
        }
        None => todo!(),
    }
    todo!()
}
